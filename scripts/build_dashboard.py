"""
FAMM - Dashboard de Gastos e Ingresos
--------------------------------------
Jala datos de la hoja "Monitoreo FAMM" en Google Drive, agrega:
  - Gastos por PROYECTO (hoja de Egresos)
  - Ingresos por Fuente de Ingreso (hoja de Ingresos)
y genera un dashboard.html estático (Chart.js) listo para publicar
en GitHub Pages.

SOLO jala las columnas necesarias (lista blanca) -- nunca RFC, nombres
de proveedores/clientes, folios fiscales, ni cualquier otro dato
sensible, aunque existan en la hoja original.

Requiere:
  - Una service account de Google con acceso de LECTURA al spreadsheet
  - Variable de entorno GOOGLE_SERVICE_ACCOUNT_JSON (contenido del
    JSON de credenciales, como secret en GitHub Actions)
  - Variable de entorno SPREADSHEET_ID
"""

import os
import json
import datetime
import gspread
from google.oauth2.service_account import Credentials

# ---------------------------------------------------------------------
# CONFIG -- AJUSTAR con los nombres reales de las pestañas del Excel
# ---------------------------------------------------------------------
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]

SHEET_EGRESOS = "Egresos"     # TODO: confirmar nombre exacto de la pestaña
SHEET_INGRESOS = "Ingresos"   # TODO: confirmar nombre exacto de la pestaña

# Lista blanca de columnas que SI se jalan de cada hoja.
# Todo lo demas (RFC, Proveedor, Cliente, Folio Fiscal, UUID CFDI,
# datos de contacto, notas, etc.) se ignora aunque exista en el Excel.
EGRESOS_COLUMNAS = {
    "proyecto": "PROYECTO",
    "tema": "TEMA",
    "programa": "PROGRAMA",
    "mes": "Mes de Pago",
    "fecha": "Fecha de Pago",
    "total": "Total",
    "status": "Status",
}

INGRESOS_COLUMNAS = {
    "fuente": "Fuente de Ingreso",
    "tipo": "Tipo de Ingreso",
    "mes": "Mes Pago",
    "fecha": "Fecha Pago",
    "total": "Total",
    "estatus": "Estatus",
}

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly",
          "https://www.googleapis.com/auth/drive.readonly"]


def get_client():
    creds_json = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    return gspread.authorize(creds)


def parse_monto(valor):
    """Convierte '$1,234.56' -> 1234.56 (float). Devuelve 0.0 si no aplica."""
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    limpio = str(valor).replace("$", "").replace(",", "").strip()
    if limpio in ("", "N/A", "n/a"):
        return 0.0
    try:
        return float(limpio)
    except ValueError:
        return 0.0


def extraer_filtrado(hoja, mapa_columnas):
    """Lee una hoja y regresa solo las columnas en mapa_columnas (lista blanca)."""
    registros = hoja.get_all_records()
    filas = []
    for r in registros:
        fila = {}
        for clave, nombre_columna in mapa_columnas.items():
            fila[clave] = r.get(nombre_columna, "")
        filas.append(fila)
    return filas


def agregar_por_categoria(filas, campo_categoria):
    """Suma 'total' agrupado por el campo indicado (proyecto o fuente)."""
    totales = {}
    for f in filas:
        categoria = (f.get(campo_categoria) or "Sin clasificar").strip() or "Sin clasificar"
        totales[categoria] = totales.get(categoria, 0.0) + parse_monto(f.get("total"))
    return dict(sorted(totales.items(), key=lambda x: x[1], reverse=True))


def generar_html(gastos_por_proyecto, ingresos_por_fuente, fecha_actualizacion):
    total_gastos = sum(gastos_por_proyecto.values())
    total_ingresos = sum(ingresos_por_fuente.values())

    def fmt(n):
        return f"${n:,.2f}"

    gastos_labels = json.dumps(list(gastos_por_proyecto.keys()))
    gastos_data = json.dumps(list(gastos_por_proyecto.values()))
    ingresos_labels = json.dumps(list(ingresos_por_fuente.keys()))
    ingresos_data = json.dumps(list(ingresos_por_fuente.values()))

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>FAMM - Dashboard Financiero</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<style>
  body {{ font-family: -apple-system, Segoe UI, Arial, sans-serif; margin: 0; padding: 24px;
         background: #f5f6f7; color: #1a1a1a; }}
  h1 {{ font-size: 20px; margin-bottom: 4px; }}
  .fecha {{ color: #666; font-size: 13px; margin-bottom: 24px; }}
  .resumen {{ display: flex; gap: 16px; margin-bottom: 32px; flex-wrap: wrap; }}
  .card {{ background: white; border-radius: 8px; padding: 16px 20px; box-shadow: 0 1px 3px rgba(0,0,0,.1); min-width: 200px; }}
  .card .label {{ font-size: 12px; color: #666; text-transform: uppercase; }}
  .card .valor {{ font-size: 24px; font-weight: 600; margin-top: 4px; }}
  .charts {{ display: flex; gap: 24px; flex-wrap: wrap; }}
  .chart-box {{ background: white; border-radius: 8px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,.1);
                flex: 1; min-width: 320px; }}
  canvas {{ max-height: 420px; }}
</style>
</head>
<body>
  <h1>FAMM - Dashboard Financiero</h1>
  <div class="fecha">Actualizado: {fecha_actualizacion}</div>

  <div class="resumen">
    <div class="card"><div class="label">Ingresos totales</div><div class="valor">{fmt(total_ingresos)}</div></div>
    <div class="card"><div class="label">Gastos totales</div><div class="valor">{fmt(total_gastos)}</div></div>
    <div class="card"><div class="label">Balance</div><div class="valor">{fmt(total_ingresos - total_gastos)}</div></div>
  </div>

  <div class="charts">
    <div class="chart-box">
      <h3>Gastos por proyecto</h3>
      <canvas id="gastosChart"></canvas>
    </div>
    <div class="chart-box">
      <h3>Ingresos por fuente</h3>
      <canvas id="ingresosChart"></canvas>
    </div>
  </div>

<script>
new Chart(document.getElementById('gastosChart'), {{
  type: 'bar',
  data: {{ labels: {gastos_labels}, datasets: [{{ label: 'Gasto (MXN)', data: {gastos_data}, backgroundColor: '#c0392b' }}] }},
  options: {{ indexAxis: 'y', plugins: {{ legend: {{ display: false }} }} }}
}});
new Chart(document.getElementById('ingresosChart'), {{
  type: 'bar',
  data: {{ labels: {ingresos_labels}, datasets: [{{ label: 'Ingreso (MXN)', data: {ingresos_data}, backgroundColor: '#27ae60' }}] }},
  options: {{ indexAxis: 'y', plugins: {{ legend: {{ display: false }} }} }}
}});
</script>
</body>
</html>
"""
    return html


def main():
    gc = get_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    hoja_egresos = sh.worksheet(SHEET_EGRESOS)
    hoja_ingresos = sh.worksheet(SHEET_INGRESOS)

    egresos = extraer_filtrado(hoja_egresos, EGRESOS_COLUMNAS)
    ingresos = extraer_filtrado(hoja_ingresos, INGRESOS_COLUMNAS)

    gastos_por_proyecto = agregar_por_categoria(egresos, "proyecto")
    ingresos_por_fuente = agregar_por_categoria(ingresos, "fuente")

    fecha_actualizacion = datetime.datetime.now().strftime("%d %b %Y, %H:%M")
    html = generar_html(gastos_por_proyecto, ingresos_por_fuente, fecha_actualizacion)

    os.makedirs("public", exist_ok=True)
    with open("public/index.html", "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Dashboard generado. {len(egresos)} egresos, {len(ingresos)} ingresos procesados.")


if __name__ == "__main__":
    main()
