# FAMM - Dashboard financiero automatizado

Jala datos de "Monitoreo FAMM 2026.xlsx" (Google Drive), genera un
dashboard de gastos por proyecto e ingresos por fuente, y lo publica
en GitHub Pages. Corre solo, sin intervención diaria.

## Qué hace

- **Todos los días (07:00 hora Monterrey):** lee las hojas de Egresos
  e Ingresos, agrega gastos por PROYECTO e ingresos por Fuente de
  Ingreso, regenera `public/index.html` y lo publica en GitHub Pages.
- **Día 15 de cada mes:** exporta el Excel original completo y lo
  sube a una carpeta de Drive de "Archivo histórico" con el nombre
  `AAAAMMDD_gastos.xlsx`.

## Privacidad

El dashboard **solo jala columnas de una lista blanca** (proyecto,
fuente, tema, programa, fecha, monto, status). Nunca toca RFC,
nombres de proveedores/clientes, folios fiscales, UUID CFDI, datos
de contacto ni ninguna otra columna del Excel, aunque exista.

## Configuración inicial (una sola vez)

1. **Crear una service account de Google** (Google Cloud Console →
   IAM → Service Accounts), generar una llave JSON.
2. **Compartir el Google Sheet** ("Monitoreo FAMM 2026") con el
   correo de la service account, con permiso de lector.
3. **Compartir la carpeta de Drive de respaldo histórico** con la
   misma service account, con permiso de editor.
4. En GitHub, ir a Settings → Secrets and variables → Actions, y
   crear:
   - `GOOGLE_SERVICE_ACCOUNT_JSON` — el contenido completo del JSON
     de la service account
   - `SPREADSHEET_ID` — el ID del spreadsheet (está en la URL,
     entre `/d/` y `/edit`)
   - `BACKUP_FOLDER_ID` — el ID de la carpeta de Drive de respaldo
5. En GitHub → Settings → Pages, elegir "GitHub Actions" como fuente.
6. **Confirmar los nombres reales de las pestañas** en
   `scripts/build_dashboard.py` (`SHEET_EGRESOS`, `SHEET_INGRESOS`) —
   quedaron como placeholder ("Egresos", "Ingresos") hasta confirmar
   el nombre exacto de cada tab en el Excel.

## Para pausar

- Apagar temporalmente: Settings → Actions → Disable workflow (uno
  o ambos). Nada se borra.
- Correr manualmente: pestaña Actions → seleccionar el workflow →
  "Run workflow".
- Eliminar todo: borrar el repositorio.
