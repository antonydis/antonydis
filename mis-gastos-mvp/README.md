# Mis gastos — MVP

Una herramienta sencilla para registrar gastos con texto, foto o voz y guardar cada movimiento en un Google Sheet del propio usuario.

## Qué hace

- Crea automáticamente `Mis gastos` en el Google Drive del usuario.
- Registra uno o varios gastos escritos en lenguaje natural.
- Lee una foto de una factura y propone el movimiento antes de guardarlo.
- Permite dictar gastos por voz desde el navegador.
- Pregunta cada mañana por los gastos del día anterior mediante correo.
- Envía un resumen semanal y mensual.
- Responde preguntas sencillas sobre los gastos sin convertir el producto en un chatbot.
- No usa una base de datos propia.
- No guarda las fotos ni los audios: se usan para extraer el movimiento y se descartan.

## Principio de privacidad

Los movimientos viven en el Google Sheet del usuario. El proyecto no tiene una base de datos central de gastos.

Para interpretar texto, imágenes y audio, el MVP usa la API de Gemini. En el nivel gratuito de Gemini, Google indica que los datos pueden usarse para mejorar sus productos. Para un despliegue con requisitos de privacidad más estrictos, usa una modalidad contractual que no utilice los datos para mejora de producto o reemplaza el proveedor por un modelo local/propio.

## Modelo

Por defecto: `gemini-3.5-flash-lite` (configurable en `Code.gs`).

El modelo solo se usa para:

1. convertir texto/foto/audio en movimientos estructurados;
2. clasificar preguntas libres en un conjunto pequeño de consultas;
3. redactar una explicación corta a partir de agregados semanales/mensuales.

Las sumas, comparaciones y agrupaciones se hacen en código, no en el modelo.

## Estructura del Sheet

- `Movimientos`: fuente de verdad de los gastos.
- `Registro diario`: permite distinguir “no gasté” de “olvidé registrar”.
- `Categorías`: lista editable por el usuario.
- `Resúmenes`: histórico semanal/mensual.
- `Configuración`: moneda, hora y zona horaria.

## Probarlo en Google Apps Script

1. Crea un proyecto nuevo en https://script.google.com/.
2. Copia `Code.gs`.
3. Crea un archivo HTML llamado `Index` y copia `Index.html`.
4. En **Configuración del proyecto**, activa “Mostrar archivo de manifiesto” y reemplaza `appsscript.json` por el incluido aquí.
5. En **Propiedades del script**, agrega `GEMINI_API_KEY` con una clave de Google AI Studio. Esto permite que la familia use el mismo procesamiento sin ver una pantalla técnica. Como alternativa, cada usuario puede guardar su propia clave desde la pantalla de activación.
6. Implementa como **Aplicación web**:
   - Ejecutar como: **Usuario que accede a la aplicación**.
   - Quién tiene acceso: **Cualquier usuario que haya iniciado sesión**.
7. Abre la URL de la implementación y autoriza los permisos.

> Importante: el modo “Usuario que accede” hace que el Sheet se cree bajo la identidad del usuario que abrió la web app. Los triggers instalables también se crean bajo esa cuenta cuando la persona completa la configuración.

## Seguridad del MVP

- La API key nunca se envía al navegador cuando se configura en Script Properties.
- Cada usuario obtiene su propio `SHEET_ID` y preferencias mediante `UserProperties`.
- No hay endpoint público para listar o leer datos de otro usuario.
- No se guardan fotos ni audios en Drive.
- El navegador comprime fotos antes de enviarlas al procesamiento.
- La interfaz siempre pide confirmación antes de guardar movimientos extraídos.

## Antes de abrirlo al público

Este MVP es adecuado para prueba familiar. Antes de convertirlo en un producto público:

- completar pantalla de privacidad y términos;
- revisar/limitar scopes OAuth;
- completar la verificación OAuth de Google cuando corresponda;
- añadir límites de tamaño, frecuencia y abuso;
- decidir si la clave de procesamiento será por usuario, por organización o mediante un backend de claves;
- revisar la política de retención del proveedor de modelos;
- añadir pruebas automáticas de extracción y cálculos;
- considerar un proveedor/modelo local para fotos y audio si la privacidad absoluta es un requisito.