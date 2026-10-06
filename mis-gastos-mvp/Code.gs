const APP = {
  NAME: 'Mis gastos',
  MODEL: 'gemini-3.5-flash-lite',
  SHEET_NAME: 'Mis gastos',
  SHEETS: {
    MOVES: 'Movimientos',
    DAYS: 'Registro diario',
    CATEGORIES: 'Categorías',
    SUMMARIES: 'Resúmenes',
    CONFIG: 'Configuración',
  },
  DEFAULT_CATEGORIES: [
    'Supermercado',
    'Restaurantes',
    'Transporte',
    'Gasolina',
    'Casa',
    'Salud',
    'Servicios',
    'Suscripciones',
    'Compras',
    'Entretenimiento',
    'Educación',
    'Viajes',
    'Familia',
    'Construcción',
    'Otros',
  ],
};

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.initialMode = (e && e.parameter && e.parameter.modo) || '';
  template.initialDate = (e && e.parameter && e.parameter.fecha) || '';
  return template
    .evaluate()
    .setTitle(APP.NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getAppState() {
  const props = PropertiesService.getUserProperties();
  const sheetId = props.getProperty('SHEET_ID');
  let ss = null;

  if (sheetId) {
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (err) {
      props.deleteProperty('SHEET_ID');
    }
  }

  const initialized = Boolean(ss);
  const config = getUserConfig_();
  const processingReady = Boolean(getGeminiApiKey_(false));

  let dashboard = null;
  if (initialized) dashboard = buildDashboard_(ss, config);

  return {
    initialized,
    processingReady,
    sheetUrl: initialized ? ss.getUrl() : '',
    config,
    dashboard,
    email: getUserEmail_(),
  };
}

function initializeUser(settings) {
  settings = settings || {};
  const props = PropertiesService.getUserProperties();
  let ss;

  const existingId = props.getProperty('SHEET_ID');
  if (existingId) {
    try {
      ss = SpreadsheetApp.openById(existingId);
    } catch (err) {
      props.deleteProperty('SHEET_ID');
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create(APP.SHEET_NAME);
    props.setProperty('SHEET_ID', ss.getId());
    createWorkbook_(ss);
  }

  const config = normalizeConfig_(settings);
  saveUserConfig_(config);
  writeConfigSheet_(ss, config);
  installUserTriggers_(config);

  return getAppState();
}

function savePersonalApiKey(apiKey) {
  const value = String(apiKey || '').trim();
  if (!value || value.length < 20) throw new Error('La clave no parece válida.');
  validateGeminiKey_(value);
  PropertiesService.getUserProperties().setProperty('GEMINI_API_KEY', value);
  return { ok: true };
}

function removePersonalApiKey() {
  PropertiesService.getUserProperties().deleteProperty('GEMINI_API_KEY');
  return { ok: true };
}

function processExpenseInput(payload) {
  ensureInitialized_();
  payload = payload || {};
  const type = String(payload.type || 'text');
  const contextDate = normalizeContextDate_(payload.contextDate);

  if (type === 'text') {
    const text = String(payload.text || '').trim();
    if (!text) throw new Error('Escribe qué gastaste.');
    return extractExpensesWithGemini_({ type: 'text', text, contextDate });
  }

  if (type === 'image' || type === 'audio') {
    const base64 = String(payload.base64 || '');
    const mimeType = String(payload.mimeType || '');
    if (!base64 || !mimeType) throw new Error('No pude leer el archivo. Intenta de nuevo.');
    return extractExpensesWithGemini_({ type, base64, mimeType, contextDate });
  }

  throw new Error('Tipo de entrada no compatible.');
}

function saveExpenses(expenses) {
  const ss = ensureInitialized_();
  if (!Array.isArray(expenses) || !expenses.length) throw new Error('No hay gastos para guardar.');

  const config = getUserConfig_();
  const sheet = ss.getSheetByName(APP.SHEETS.MOVES);
  const categories = getCategories_(ss);
  const rows = [];
  const touchedDates = {};
  const now = new Date();

  expenses.forEach((item) => {
    const clean = validateExpense_(item, config, categories);
    touchedDates[clean.date] = true;
    rows.push([
      Utilities.getUuid(),
      clean.date,
      clean.description,
      clean.merchant,
      clean.category,
      clean.amount,
      clean.currency,
      clean.paymentMethod,
      clean.source,
      now,
    ]);
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  Object.keys(touchedDates).forEach((date) => upsertDailyStatus_(ss, date, 'Registrado'));

  return {
    ok: true,
    count: rows.length,
    dashboard: buildDashboard_(ss, config),
  };
}

function markNoExpenses(dateText) {
  const ss = ensureInitialized_();
  const date = normalizeContextDate_(dateText || 'yesterday');
  upsertDailyStatus_(ss, date, 'Sin gastos');
  return { ok: true, date, dashboard: buildDashboard_(ss, getUserConfig_()) };
}

function getDashboard() {
  const ss = ensureInitialized_();
  return buildDashboard_(ss, getUserConfig_());
}

function runQuickQuestion(type, category) {
  const ss = ensureInitialized_();
  const config = getUserConfig_();
  const rows = readTransactions_(ss);
  return answerIntent_(type, category, rows, config);
}

function askQuestion(question) {
  const ss = ensureInitialized_();
  const config = getUserConfig_();
  const text = String(question || '').trim();
  if (!text) throw new Error('Escribe una pregunta.');

  const categories = getCategories_(ss);
  const intent = classifyQuestion_(text, categories);
  return answerIntent_(intent.intent, intent.category || '', readTransactions_(ss), config);
}

function sendDailyReminder() {
  const ss = getUserSpreadsheet_();
  if (!ss) return;
  const config = getUserConfig_();
  const yesterday = formatDateInTz_(addDays_(new Date(), -1), config.timeZone);
  if (getDailyStatus_(ss, yesterday)) return;

  const email = getUserEmail_();
  if (!email) return;
  const url = ScriptApp.getService().getUrl();
  if (!url) return;

  const registerUrl = url + '?modo=registro&fecha=' + encodeURIComponent(yesterday);
  const noSpendUrl = url + '?modo=sin-gastos&fecha=' + encodeURIComponent(yesterday);

  const html = [
    '<p>Buenos días.</p>',
    '<p><strong>¿Qué gastaste ayer?</strong></p>',
    '<p><a href="' + registerUrl + '">Registrar gastos</a> &nbsp;·&nbsp; <a href="' + noSpendUrl + '">No gasté nada</a></p>',
    '<p style="color:#666;font-size:12px">Tus movimientos se guardan en tu propia hoja de Google.</p>',
  ].join('');

  MailApp.sendEmail({
    to: email,
    subject: '¿Qué gastaste ayer?',
    htmlBody: html,
    body: '¿Qué gastaste ayer?\n\nRegistrar: ' + registerUrl + '\nNo gasté nada: ' + noSpendUrl,
    name: APP.NAME,
  });
}

function sendWeeklySummary() {
  const ss = getUserSpreadsheet_();
  if (!ss) return;
  const config = getUserConfig_();
  const email = getUserEmail_();
  if (!email) return;

  const end = startOfDay_(addDays_(new Date(), -1));
  const start = addDays_(end, -6);
  const summary = createPeriodSummary_(ss, start, end, 'Semana', config);
  sendSummaryEmail_(email, 'Tu semana en gastos', summary);
}

function sendMonthlySummary() {
  const ss = getUserSpreadsheet_();
  if (!ss) return;
  const config = getUserConfig_();
  const email = getUserEmail_();
  if (!email) return;

  const now = new Date();
  const firstThisMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const end = addDays_(firstThisMonth, -1);
  const start = new Date(end.getFullYear(), end.getMonth(), 1);
  const summary = createPeriodSummary_(ss, start, end, 'Mes', config);
  sendSummaryEmail_(email, 'Tu mes en gastos', summary);
}

function createWorkbook_(ss) {
  const first = ss.getSheets()[0];
  first.setName(APP.SHEETS.MOVES);
  setupSheet_(first, ['ID', 'Fecha', 'Descripción', 'Comercio', 'Categoría', 'Monto', 'Moneda', 'Método', 'Fuente', 'Registrado']);
  first.setFrozenRows(1);
  first.getRange('F:F').setNumberFormat('#,##0.00');
  first.getRange('J:J').setNumberFormat('yyyy-mm-dd hh:mm');

  const days = ss.insertSheet(APP.SHEETS.DAYS);
  setupSheet_(days, ['Fecha', 'Estado', 'Registrado']);
  days.setFrozenRows(1);

  const categories = ss.insertSheet(APP.SHEETS.CATEGORIES);
  setupSheet_(categories, ['Categoría', 'Activa', 'Palabras clave']);
  categories.getRange(2, 1, APP.DEFAULT_CATEGORIES.length, 3).setValues(
    APP.DEFAULT_CATEGORIES.map((name) => [name, true, ''])
  );
  categories.setFrozenRows(1);

  const summaries = ss.insertSheet(APP.SHEETS.SUMMARIES);
  setupSheet_(summaries, ['Periodo', 'Desde', 'Hasta', 'Total', 'Moneda', 'Mayor categoría', '% mayor categoría', 'Variación %', 'Resumen', 'Generado']);
  summaries.setFrozenRows(1);

  const config = ss.insertSheet(APP.SHEETS.CONFIG);
  setupSheet_(config, ['Ajuste', 'Valor']);
  config.setFrozenRows(1);

  ss.setActiveSheet(first);
}

function setupSheet_(sheet, headers) {
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
}

function writeConfigSheet_(ss, config) {
  const sheet = ss.getSheetByName(APP.SHEETS.CONFIG);
  const values = [
    ['Ajuste', 'Valor'],
    ['Moneda principal', config.currency],
    ['Hora del recordatorio', config.reminderHour],
    ['Zona horaria', config.timeZone],
    ['Correo', getUserEmail_()],
  ];
  sheet.clear();
  sheet.getRange(1, 1, values.length, 2).setValues(values);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  sheet.autoResizeColumns(1, 2);
}

function installUserTriggers_(config) {
  const handlers = ['sendDailyReminder', 'sendWeeklySummary', 'sendMonthlySummary'];
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (handlers.indexOf(trigger.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(trigger);
  });

  ScriptApp.newTrigger('sendDailyReminder')
    .timeBased()
    .atHour(Number(config.reminderHour))
    .everyDays(1)
    .inTimezone(config.timeZone)
    .create();

  ScriptApp.newTrigger('sendWeeklySummary')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(Math.min(Number(config.reminderHour) + 1, 23))
    .inTimezone(config.timeZone)
    .create();

  ScriptApp.newTrigger('sendMonthlySummary')
    .timeBased()
    .onMonthDay(1)
    .atHour(Math.min(Number(config.reminderHour) + 1, 23))
    .inTimezone(config.timeZone)
    .create();
}

function normalizeConfig_(settings) {
  const old = getUserConfig_();
  const reminderHour = Number(settings.reminderHour != null ? settings.reminderHour : old.reminderHour);
  return {
    currency: String(settings.currency || old.currency || 'NIO').toUpperCase().slice(0, 3),
    reminderHour: Math.max(0, Math.min(23, isFinite(reminderHour) ? reminderHour : 8)),
    timeZone: String(settings.timeZone || old.timeZone || Session.getScriptTimeZone() || 'America/Managua'),
  };
}

function saveUserConfig_(config) {
  PropertiesService.getUserProperties().setProperties({
    CURRENCY: config.currency,
    REMINDER_HOUR: String(config.reminderHour),
    TIME_ZONE: config.timeZone,
    USER_EMAIL: getUserEmail_(),
  });
}

function getUserConfig_() {
  const p = PropertiesService.getUserProperties();
  return {
    currency: p.getProperty('CURRENCY') || 'NIO',
    reminderHour: Number(p.getProperty('REMINDER_HOUR') || 8),
    timeZone: p.getProperty('TIME_ZONE') || Session.getScriptTimeZone() || 'America/Managua',
  };
}

function getUserEmail_() {
  const p = PropertiesService.getUserProperties();
  return p.getProperty('USER_EMAIL') || Session.getEffectiveUser().getEmail() || '';
}

function getUserSpreadsheet_() {
  const id = PropertiesService.getUserProperties().getProperty('SHEET_ID');
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch (err) { return null; }
}

function ensureInitialized_() {
  const ss = getUserSpreadsheet_();
  if (!ss) throw new Error('Primero crea tu registro.');
  return ss;
}

function getGeminiApiKey_(required) {
  const userKey = PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
  const sharedKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const key = userKey || sharedKey || '';
  if (!key && required !== false) throw new Error('Falta activar la lectura de texto, fotos y voz.');
  return key;
}

function validateGeminiKey_(apiKey) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(APP.MODEL) + ':generateContent';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify({ contents: [{ parts: [{ text: 'Responde únicamente: OK' }] }] }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) throw new Error('No pude validar la clave.');
}

function extractExpensesWithGemini_(input) {
  const ss = ensureInitialized_();
  const config = getUserConfig_();
  const categories = getCategories_(ss);
  const apiKey = getGeminiApiKey_(true);
  const today = formatDateInTz_(new Date(), config.timeZone);

  const prompt = [
    'Tu única tarea es convertir una entrada de gastos personales en datos estructurados.',
    'Hoy es ' + today + '. La fecha sugerida para esta entrada es ' + input.contextDate + '.',
    'Moneda principal: ' + config.currency + '.',
    'Categorías permitidas: ' + categories.join(', ') + '.',
    '',
    'Reglas:',
    '- No inventes montos, fechas ni comercios.',
    '- Si el usuario da varios gastos, devuelve varios movimientos.',
    '- Si es una foto de factura, registra el TOTAL pagado como un movimiento, no cada producto.',
    '- Si no se menciona moneda, usa la moneda principal.',
    '- Usa la categoría permitida más cercana. Usa Otros solo cuando realmente no haya una categoría mejor.',
    '- Método de pago puede quedar vacío y no requiere pregunta.',
    '- Pregunta solo cuando falte un dato esencial o haya una ambigüedad real que cambie el registro, especialmente monto, fecha o una categoría muy dudosa.',
    '- Si puedes resolverlo razonablemente sin preguntar, hazlo.',
    '- Descripción corta, natural y en español.',
    '- source debe ser texto, foto o voz según la entrada.',
  ].join('\n');

  const parts = [{ text: prompt }];
  if (input.type === 'text') {
    parts.push({ text: 'Entrada del usuario:\n' + input.text });
  } else {
    parts.push({
      inlineData: {
        mimeType: input.mimeType,
        data: input.base64,
      },
    });
    parts.push({ text: 'Analiza este ' + (input.type === 'image' ? 'comprobante o imagen' : 'audio') + ' como registro de gastos.' });
  }

  const schema = {
    type: 'object',
    properties: {
      needs_clarification: { type: 'boolean' },
      question: { type: 'string' },
      expenses: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            date: { type: 'string', description: 'YYYY-MM-DD' },
            description: { type: 'string' },
            merchant: { type: 'string' },
            category: { type: 'string', enum: categories },
            amount: { type: 'number' },
            currency: { type: 'string' },
            payment_method: { type: 'string' },
          },
          required: ['date', 'description', 'merchant', 'category', 'amount', 'currency', 'payment_method'],
        },
      },
    },
    required: ['needs_clarification', 'question', 'expenses'],
  };

  const body = {
    contents: [{ role: 'user', parts }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json',
      responseSchema: schema,
    },
  };

  const data = callGemini_(apiKey, body);
  const parsed = parseGeminiJson_(data);
  const result = {
    needsClarification: Boolean(parsed.needs_clarification),
    question: String(parsed.question || ''),
    expenses: Array.isArray(parsed.expenses) ? parsed.expenses.map((x) => ({
      date: normalizeDateString_(x.date, input.contextDate),
      description: String(x.description || 'Gasto').slice(0, 120),
      merchant: String(x.merchant || '').slice(0, 120),
      category: categories.indexOf(x.category) >= 0 ? x.category : 'Otros',
      amount: Number(x.amount || 0),
      currency: String(x.currency || config.currency).toUpperCase().slice(0, 3),
      paymentMethod: String(x.payment_method || '').slice(0, 60),
      source: input.type === 'image' ? 'Foto' : input.type === 'audio' ? 'Voz' : 'Texto',
    })).filter((x) => x.amount > 0) : [],
  };

  if (!result.expenses.length && !result.needsClarification) {
    result.needsClarification = true;
    result.question = 'No pude identificar un monto con claridad. ¿Cuánto gastaste?';
  }
  return result;
}

function classifyQuestion_(question, categories) {
  const apiKey = getGeminiApiKey_(false);
  if (!apiKey) return classifyQuestionLocally_(question, categories);

  const schema = {
    type: 'object',
    properties: {
      intent: {
        type: 'string',
        enum: ['month_total', 'top_category', 'compare_last_month', 'largest_expenses', 'category_total', 'review_spending', 'unknown'],
      },
      category: { type: 'string' },
    },
    required: ['intent', 'category'],
  };

  const body = {
    contents: [{ parts: [{ text: [
      'Clasifica esta pregunta sobre gastos personales.',
      'Categorías disponibles: ' + categories.join(', ') + '.',
      'No respondas la pregunta, solo clasifícala.',
      'Pregunta: ' + question,
    ].join('\n') }] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: schema },
  };

  try {
    const parsed = parseGeminiJson_(callGemini_(apiKey, body));
    return { intent: parsed.intent || 'unknown', category: parsed.category || '' };
  } catch (err) {
    return classifyQuestionLocally_(question, categories);
  }
}

function classifyQuestionLocally_(question, categories) {
  const q = String(question || '').toLowerCase();
  if (/mes pasado|compar|más o menos/.test(q)) return { intent: 'compare_last_month', category: '' };
  if (/categoría|en qué.*más|dónde.*más/.test(q)) return { intent: 'top_category', category: '' };
  if (/grande|mayor(es)? gasto/.test(q)) return { intent: 'largest_expenses', category: '' };
  if (/revis|algo.*mirar|patr[oó]n|llama.*atenci[oó]n/.test(q)) return { intent: 'review_spending', category: '' };
  for (let i = 0; i < categories.length; i++) {
    if (q.indexOf(categories[i].toLowerCase()) >= 0) return { intent: 'category_total', category: categories[i] };
  }
  if (/mes|total|cu[aá]nto.*gast/.test(q)) return { intent: 'month_total', category: '' };
  return { intent: 'unknown', category: '' };
}

function answerIntent_(type, category, rows, config) {
  const now = new Date();
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const prevMonthStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const current = filterRowsByDate_(rows, monthStart, nextMonth);
  const previous = filterRowsByDate_(rows, prevMonthStart, monthStart);

  if (type === 'month_total') {
    const total = sumRows_(current);
    return { answer: 'Este mes llevas ' + money_(total, config.currency) + '.', type };
  }

  if (type === 'top_category') {
    const grouped = groupByCategory_(current);
    const top = firstEntry_(grouped);
    if (!top) return { answer: 'Todavía no tienes gastos registrados este mes.', type };
    return { answer: 'Este mes estás gastando más en ' + top.name + ': ' + money_(top.value, config.currency) + '.', type };
  }

  if (type === 'compare_last_month') {
    const a = sumRows_(current);
    const b = sumRows_(previous);
    if (!b) return { answer: 'No tengo suficiente información del mes pasado para comparar todavía.', type };
    const pct = ((a - b) / b) * 100;
    const direction = pct >= 0 ? 'más' : 'menos';
    return { answer: 'Llevas ' + money_(a, config.currency) + ' este mes, ' + Math.abs(pct).toFixed(0) + '% ' + direction + ' que el mes pasado completo (' + money_(b, config.currency) + ').', type };
  }

  if (type === 'largest_expenses') {
    const biggest = current.slice().sort((a, b) => b.amount - a.amount).slice(0, 5);
    if (!biggest.length) return { answer: 'Todavía no tienes gastos registrados este mes.', type };
    return {
      answer: 'Tus gastos más grandes este mes son:\n' + biggest.map((r, i) => (i + 1) + '. ' + r.description + ': ' + money_(r.amount, r.currency || config.currency)).join('\n'),
      type,
    };
  }

  if (type === 'category_total') {
    const wanted = findClosestCategory_(category, getCategoriesFromRows_(rows));
    if (!wanted) return { answer: 'Dime la categoría que quieres revisar.', type };
    const selected = current.filter((r) => r.category === wanted);
    return { answer: 'Este mes gastaste ' + money_(sumRows_(selected), config.currency) + ' en ' + wanted + '.', type };
  }

  if (type === 'review_spending') {
    return { answer: buildSpendingReview_(rows, config), type };
  }

  return {
    answer: 'Puedo decirte cuánto llevas este mes, en qué categoría gastas más, compararlo con el mes pasado, mostrar tus gastos más grandes o revisar una categoría.',
    type: 'unknown',
  };
}

function buildSpendingReview_(rows, config) {
  const now = new Date();
  const end = addDays_(startOfDay_(now), 1);
  const start = addDays_(end, -30);
  const current = filterRowsByDate_(rows, start, end);
  if (!current.length) return 'Todavía no tengo suficientes gastos para detectar algo útil.';

  const total = sumRows_(current);
  const grouped = groupByCategory_(current);
  const top = firstEntry_(grouped);
  const biggest = current.slice().sort((a, b) => b.amount - a.amount)[0];
  const facts = {
    total30Days: total,
    currency: config.currency,
    topCategory: top ? top.name : '',
    topCategoryAmount: top ? top.value : 0,
    topCategoryShare: top && total ? Math.round((top.value / total) * 100) : 0,
    largestExpense: biggest ? { description: biggest.description, amount: biggest.amount, category: biggest.category } : null,
    categories: Object.keys(grouped).slice(0, 8).map((name) => ({ name, amount: grouped[name] })),
  };

  const apiKey = getGeminiApiKey_(false);
  if (!apiKey) {
    return 'En los últimos 30 días gastaste ' + money_(total, config.currency) + '. Tu categoría principal fue ' + (top ? top.name : 'sin datos') + (top ? ' con ' + Math.round((top.value / total) * 100) + '% del total.' : '.');
  }

  const prompt = [
    'Explica estos datos de gastos personales en español claro y cotidiano.',
    'Máximo 3 frases. No des sermones ni consejos financieros genéricos.',
    'Señala lo más útil o llamativo. No inventes datos.',
    JSON.stringify(facts),
  ].join('\n');

  try {
    const data = callGemini_(apiKey, { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2 } });
    return getGeminiText_(data).trim();
  } catch (err) {
    return 'En los últimos 30 días gastaste ' + money_(total, config.currency) + '. Tu categoría principal fue ' + (top ? top.name : 'sin datos') + '.';
  }
}

function createPeriodSummary_(ss, start, end, label, config) {
  const rows = readTransactions_(ss);
  const periodRows = filterRowsByDate_(rows, start, addDays_(end, 1));
  const previousStart = addDays_(start, -daysBetween_(start, end) - 1);
  const previousEndExclusive = start;
  const previousRows = filterRowsByDate_(rows, previousStart, previousEndExclusive);
  const total = sumRows_(periodRows);
  const prev = sumRows_(previousRows);
  const grouped = groupByCategory_(periodRows);
  const top = firstEntry_(grouped);
  const variation = prev ? ((total - prev) / prev) * 100 : null;

  const facts = {
    period: label,
    from: formatDateInTz_(start, config.timeZone),
    to: formatDateInTz_(end, config.timeZone),
    total,
    currency: config.currency,
    previousTotal: prev,
    variationPercent: variation == null ? null : Math.round(variation),
    topCategories: Object.keys(grouped).slice(0, 5).map((name) => ({ name, amount: grouped[name] })),
  };

  let narrative = total ? buildSummaryNarrative_(facts) : 'No registraste gastos en este período.';
  const summarySheet = ss.getSheetByName(APP.SHEETS.SUMMARIES);
  summarySheet.appendRow([
    label,
    facts.from,
    facts.to,
    total,
    config.currency,
    top ? top.name : '',
    top && total ? Math.round((top.value / total) * 100) : '',
    variation == null ? '' : Math.round(variation),
    narrative,
    new Date(),
  ]);

  return { facts, narrative };
}

function buildSummaryNarrative_(facts) {
  const apiKey = getGeminiApiKey_(false);
  if (!apiKey) {
    const top = facts.topCategories[0];
    let text = 'Gastaste ' + money_(facts.total, facts.currency) + '.';
    if (top) text += ' La categoría principal fue ' + top.name + ' con ' + money_(top.amount, facts.currency) + '.';
    if (facts.variationPercent != null) text += ' Eso es ' + Math.abs(facts.variationPercent) + '% ' + (facts.variationPercent >= 0 ? 'más' : 'menos') + ' que el período anterior.';
    return text;
  }

  const prompt = [
    'Redacta un resumen corto de gastos personales en español.',
    'Máximo 4 frases. Tono claro, neutral y útil. Sin mencionar IA.',
    'No inventes. No des recomendaciones si los datos no las respaldan.',
    JSON.stringify(facts),
  ].join('\n');

  try {
    return getGeminiText_(callGemini_(apiKey, { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2 } })).trim();
  } catch (err) {
    return 'Gastaste ' + money_(facts.total, facts.currency) + ' en este período.';
  }
}

function sendSummaryEmail_(email, subject, summary) {
  const facts = summary.facts;
  const categories = (facts.topCategories || []).map((x) => '<li>' + escapeHtml_(x.name) + ': <strong>' + escapeHtml_(money_(x.amount, facts.currency)) + '</strong></li>').join('');
  const url = ScriptApp.getService().getUrl() || '';
  const html = [
    '<p><strong>' + escapeHtml_(money_(facts.total, facts.currency)) + '</strong> en total.</p>',
    '<p>' + escapeHtml_(summary.narrative) + '</p>',
    categories ? '<ul>' + categories + '</ul>' : '',
    url ? '<p><a href="' + url + '">Ver mis gastos</a></p>' : '',
  ].join('');

  MailApp.sendEmail({
    to: email,
    subject,
    htmlBody: html,
    body: summary.narrative + '\n\nTotal: ' + money_(facts.total, facts.currency) + (url ? '\n\nVer: ' + url : ''),
    name: APP.NAME,
  });
}

function callGemini_(apiKey, body) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(APP.MODEL) + ':generateContent';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  let data;
  try { data = JSON.parse(res.getContentText()); } catch (err) { data = null; }
  if (code < 200 || code >= 300) {
    const msg = data && data.error && data.error.message ? data.error.message : 'No pude procesar la entrada.';
    throw new Error(msg);
  }
  return data;
}

function getGeminiText_(data) {
  const candidates = data && data.candidates;
  const parts = candidates && candidates[0] && candidates[0].content && candidates[0].content.parts;
  if (!parts || !parts.length) throw new Error('No recibí una respuesta válida.');
  return parts.map((p) => p.text || '').join('');
}

function parseGeminiJson_(data) {
  const text = getGeminiText_(data).trim().replace(/^```json\s*/i, '').replace(/```$/i, '').trim();
  return JSON.parse(text);
}

function getCategories_(ss) {
  const sheet = ss.getSheetByName(APP.SHEETS.CATEGORIES);
  if (!sheet || sheet.getLastRow() < 2) return APP.DEFAULT_CATEGORIES.slice();
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const active = values.filter((r) => r[0] && r[1] !== false).map((r) => String(r[0]).trim()).filter(Boolean);
  return active.length ? active : APP.DEFAULT_CATEGORIES.slice();
}

function validateExpense_(item, config, categories) {
  const amount = Number(item.amount);
  if (!isFinite(amount) || amount <= 0) throw new Error('Hay un monto inválido.');
  const date = normalizeDateString_(item.date, formatDateInTz_(new Date(), config.timeZone));
  return {
    date,
    description: String(item.description || 'Gasto').trim().slice(0, 120),
    merchant: String(item.merchant || '').trim().slice(0, 120),
    category: categories.indexOf(item.category) >= 0 ? item.category : 'Otros',
    amount: Math.round(amount * 100) / 100,
    currency: String(item.currency || config.currency).trim().toUpperCase().slice(0, 3),
    paymentMethod: String(item.paymentMethod || item.payment_method || '').trim().slice(0, 60),
    source: String(item.source || 'Texto').trim().slice(0, 20),
  };
}

function upsertDailyStatus_(ss, date, status) {
  const sheet = ss.getSheetByName(APP.SHEETS.DAYS);
  const last = sheet.getLastRow();
  if (last >= 2) {
    const values = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === date) {
        sheet.getRange(i + 2, 2, 1, 2).setValues([[status, new Date()]]);
        return;
      }
    }
  }
  sheet.appendRow([date, status, new Date()]);
}

function getDailyStatus_(ss, date) {
  const sheet = ss.getSheetByName(APP.SHEETS.DAYS);
  if (!sheet || sheet.getLastRow() < 2) return '';
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
  for (let i = 0; i < values.length; i++) if (values[i][0] === date) return values[i][1];
  return '';
}

function readTransactions_(ss) {
  const sheet = ss.getSheetByName(APP.SHEETS.MOVES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  return values.map((r) => ({
    id: r[0],
    date: normalizeSheetDate_(r[1]),
    description: String(r[2] || ''),
    merchant: String(r[3] || ''),
    category: String(r[4] || 'Otros'),
    amount: Number(r[5] || 0),
    currency: String(r[6] || ''),
    paymentMethod: String(r[7] || ''),
    source: String(r[8] || ''),
  })).filter((r) => r.amount > 0 && r.date);
}

function buildDashboard_(ss, config) {
  const rows = readTransactions_(ss);
  const now = new Date();
  const today = formatDateInTz_(now, config.timeZone);
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const month = filterRowsByDate_(rows, monthStart, nextMonth);
  const grouped = groupByCategory_(month);
  const top = firstEntry_(grouped);
  const latest = rows.slice().sort((a, b) => b.date.localeCompare(a.date)).slice(0, 5);

  return {
    today,
    monthTotal: sumRows_(month),
    currency: config.currency,
    topCategory: top ? top.name : '',
    topCategoryAmount: top ? top.value : 0,
    latest: latest.map((r) => ({ date: r.date, description: r.description, category: r.category, amount: r.amount, currency: r.currency || config.currency })),
  };
}

function filterRowsByDate_(rows, start, endExclusive) {
  const startKey = dateKeyLocal_(start);
  const endKey = dateKeyLocal_(endExclusive);
  return rows.filter((r) => r.date >= startKey && r.date < endKey);
}

function sumRows_(rows) {
  return rows.reduce((sum, r) => sum + Number(r.amount || 0), 0);
}

function groupByCategory_(rows) {
  const raw = {};
  rows.forEach((r) => { raw[r.category || 'Otros'] = (raw[r.category || 'Otros'] || 0) + Number(r.amount || 0); });
  return Object.keys(raw).sort((a, b) => raw[b] - raw[a]).reduce((acc, k) => { acc[k] = raw[k]; return acc; }, {});
}

function firstEntry_(obj) {
  const keys = Object.keys(obj || {});
  return keys.length ? { name: keys[0], value: obj[keys[0]] } : null;
}

function getCategoriesFromRows_(rows) {
  const seen = {};
  rows.forEach((r) => { if (r.category) seen[r.category] = true; });
  return Object.keys(seen);
}

function findClosestCategory_(text, categories) {
  const q = String(text || '').trim().toLowerCase();
  if (!q) return '';
  for (let i = 0; i < categories.length; i++) if (categories[i].toLowerCase() === q) return categories[i];
  for (let i = 0; i < categories.length; i++) if (categories[i].toLowerCase().indexOf(q) >= 0 || q.indexOf(categories[i].toLowerCase()) >= 0) return categories[i];
  return '';
}

function normalizeContextDate_(value) {
  const config = getUserConfig_();
  const today = new Date();
  if (!value || value === 'today' || value === 'hoy') return formatDateInTz_(today, config.timeZone);
  if (value === 'yesterday' || value === 'ayer') return formatDateInTz_(addDays_(today, -1), config.timeZone);
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(value))) return String(value);
  return formatDateInTz_(today, config.timeZone);
}

function normalizeDateString_(value, fallback) {
  const text = String(value || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text : fallback;
}

function normalizeSheetDate_(value) {
  if (value instanceof Date && !isNaN(value)) return dateKeyLocal_(value);
  const text = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  return '';
}

function money_(amount, currency) {
  const n = Number(amount || 0);
  return String(currency || '') + ' ' + n.toLocaleString('es', { minimumFractionDigits: n % 1 ? 2 : 0, maximumFractionDigits: 2 });
}

function formatDateInTz_(date, timeZone) {
  return Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
}

function dateKeyLocal_(date) {
  return Utilities.formatDate(date, getUserConfig_().timeZone, 'yyyy-MM-dd');
}

function addDays_(date, days) {
  const d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function daysBetween_(start, end) {
  return Math.round((startOfDay_(end).getTime() - startOfDay_(start).getTime()) / 86400000);
}

function escapeHtml_(text) {
  return String(text || '').replace(/[&<>'"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[c]));
}