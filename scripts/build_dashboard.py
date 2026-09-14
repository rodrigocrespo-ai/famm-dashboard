"""
FAMM - Dashboard de Gastos e Ingresos (v2)
--------------------------------------------
Jala el archivo "Monitoreo FAMM 2026.xlsx" de Google Drive y genera un
dashboard.html con:
  - Selector de año (compara meses entre distintos años)
  - Gastos por proyecto (excluye traspasos entre cuentas; proyectos que
    individualmente pesan menos del 2% se agrupan en "Otros")
  - Ingresos: dos tarjetas -- Compensaciones Ambientales vs
    Asociados y Proyectos (segun columna "Programa / Proyecto")
  - Ingresos por fuente (excluye traspasos entre cuentas)
  - Comparativo mensual de Ingresos y Gastos, una linea por año

SOLO jala columnas de una lista blanca -- nunca RFC, nombres de
proveedores/clientes, folios fiscales, ni cualquier otro dato sensible,
aunque existan en el Excel.

Requiere:
  - Service account de Google con acceso de LECTURA al archivo
  - GOOGLE_SERVICE_ACCOUNT_JSON (secret en GitHub Actions)
  - SPREADSHEET_ID
"""

import os
import io
import json
import datetime
import openpyxl
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]

SHEET_EGRESOS = "Egreso"
SHEET_INGRESOS = "Ingreso"

UMBRAL_OTROS = 0.02  # proyectos con menos del 2% del total se agrupan en "Otros"
TEMA_TRASPASO = "9 Traspaso entre cuentas"
TIPO_INGRESO_TRASPASO = "Traspaso entre cuentas"
PROGRAMA_COMPENSACION = "Compensación"

# Lista blanca de columnas (por encabezado exacto en la fila 1 del Excel).
EGRESOS_COLUMNAS = {
    "proyecto": "PROYECTO",
    "tema": "TEMA",
    "programa": "PROGRAMA",
    "fecha": "Fecha de Pago",
    "total": "Total",
    "status": "Status",
}

INGRESOS_COLUMNAS = {
    "fuente": "Fuente de Ingreso",
    "tipo": "Tipo de Ingreso",
    "programa_proyecto": "Programa / Proyecto",
    "fecha": "Fecha Pago",
    "total": "Total",
    "estatus": "Estatus",
}

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

MESES_ES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
            "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def get_drive_service():
    creds_json = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def descargar_excel(service):
    request = service.files().get_media(fileId=SPREADSHEET_ID)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buffer.seek(0)
    return buffer


def parse_monto(valor):
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


def parse_fecha(valor):
    """Regresa (año, mes) o (None, None) si no se puede interpretar."""
    if valor is None:
        return None, None
    if isinstance(valor, (datetime.datetime, datetime.date)):
        return valor.year, valor.month
    texto = str(valor).strip()
    for fmt in ("%d %b %Y", "%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            d = datetime.datetime.strptime(texto, fmt)
            return d.year, d.month
        except ValueError:
            continue
    return None, None


def extraer_filtrado(hoja, mapa_columnas):
    filas_iter = hoja.iter_rows(values_only=True)
    encabezados = next(filas_iter)
    indice_columna = {}
    for idx, nombre in enumerate(encabezados):
        if nombre is None:
            continue
        indice_columna[str(nombre).strip()] = idx

    filas = []
    for row in filas_iter:
        if row is None or all(v is None for v in row):
            continue
        fila = {}
        for clave, nombre_columna in mapa_columnas.items():
            idx = indice_columna.get(nombre_columna)
            fila[clave] = row[idx] if idx is not None and idx < len(row) else ""
        anio, mes = parse_fecha(fila.get("fecha"))
        fila["anio"] = anio
        fila["mes"] = mes
        fila["monto"] = parse_monto(fila.get("total"))
        filas.append(fila)
    return filas


def aplicar_umbral_otros(totales, umbral=UMBRAL_OTROS):
    """Agrupa categorias que pesan menos de `umbral` del total en 'Otros'."""
    total_general = sum(totales.values())
    if total_general == 0:
        return totales
    principales = {}
    otros = 0.0
    for categoria, monto in totales.items():
        if monto / total_general < umbral:
            otros += monto
        else:
            principales[categoria] = monto
    if otros > 0:
        principales["Otros"] = otros
    return dict(sorted(principales.items(), key=lambda x: x[1], reverse=True))


def main():
    service = get_drive_service()
    excel_bytes = descargar_excel(service)
    wb = openpyxl.load_workbook(excel_bytes, data_only=True, read_only=True)

    egresos_raw = extraer_filtrado(wb[SHEET_EGRESOS], EGRESOS_COLUMNAS)
    ingresos_raw = extraer_filtrado(wb[SHEET_INGRESOS], INGRESOS_COLUMNAS)

    # Excluir traspasos entre cuentas de ambos lados
    egresos = [f for f in egresos_raw if str(f.get("tema", "")).strip() != TEMA_TRASPASO]
    ingresos = [f for f in ingresos_raw if str(f.get("tipo", "")).strip() != TIPO_INGRESO_TRASPASO]

    anios = sorted({f["anio"] for f in egresos + ingresos if f["anio"]}, reverse=True)

    gastos_por_proyecto_anio = {}
    ingresos_por_fuente_anio = {}
    ingresos_compensacion_anio = {}
    ingresos_asociados_anio = {}
    monthly_ingresos = {}
    monthly_gastos = {}

    for anio in anios:
        egresos_anio = [f for f in egresos if f["anio"] == anio]
        ingresos_anio = [f for f in ingresos if f["anio"] == anio]

        # Gastos por proyecto, con umbral de "Otros"
        totales_proyecto = {}
        for f in egresos_anio:
            proyecto = str(f.get("proyecto") or "Sin clasificar").strip() or "Sin clasificar"
            totales_proyecto[proyecto] = totales_proyecto.get(proyecto, 0.0) + f["monto"]
        gastos_por_proyecto_anio[anio] = aplicar_umbral_otros(totales_proyecto)

        # Ingresos por fuente
        totales_fuente = {}
        for f in ingresos_anio:
            fuente = str(f.get("fuente") or "Sin clasificar").strip() or "Sin clasificar"
            totales_fuente[fuente] = totales_fuente.get(fuente, 0.0) + f["monto"]
        ingresos_por_fuente_anio[anio] = dict(sorted(totales_fuente.items(), key=lambda x: x[1], reverse=True))

        # Ingresos: Compensaciones Ambientales vs Asociados y Proyectos
        comp = sum(f["monto"] for f in ingresos_anio
                   if str(f.get("programa_proyecto", "")).strip() == PROGRAMA_COMPENSACION)
        asoc = sum(f["monto"] for f in ingresos_anio
                   if str(f.get("programa_proyecto", "")).strip() != PROGRAMA_COMPENSACION)
        ingresos_compensacion_anio[anio] = comp
        ingresos_asociados_anio[anio] = asoc

        # Comparativo mensual
        meses_ingresos = [0.0] * 12
        meses_gastos = [0.0] * 12
        for f in ingresos_anio:
            if f["mes"]:
                meses_ingresos[f["mes"] - 1] += f["monto"]
        for f in egresos_anio:
            if f["mes"]:
                meses_gastos[f["mes"] - 1] += f["monto"]
        monthly_ingresos[anio] = meses_ingresos
        monthly_gastos[anio] = meses_gastos

    fecha_actualizacion = datetime.datetime.now().strftime("%d %b %Y, %H:%M")

    data_js = {
        "anios": anios,
        "gastos_por_proyecto": gastos_por_proyecto_anio,
        "ingresos_por_fuente": ingresos_por_fuente_anio,
        "ingresos_compensacion": ingresos_compensacion_anio,
        "ingresos_asociados": ingresos_asociados_anio,
        "monthly_ingresos": monthly_ingresos,
        "monthly_gastos": monthly_gastos,
        "meses_es": MESES_ES,
    }

    html = generar_html(data_js, fecha_actualizacion)

    os.makedirs("public", exist_ok=True)
    with open("public/index.html", "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Dashboard generado. Años: {anios}. "
          f"{len(egresos)} egresos, {len(ingresos)} ingresos procesados (traspasos excluidos).")


def generar_html(data, fecha_actualizacion):
    data_json = json.dumps(data, ensure_ascii=False)

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
  h1 {{ font-size: 20px; margin-bottom: 4px; display: inline-block; }}
  .header {{ display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; }}
  .fecha {{ color: #666; font-size: 13px; margin-bottom: 24px; }}
  select {{ font-size: 15px; padding: 6px 10px; border-radius: 6px; border: 1px solid #ccc; background: white; }}
  .resumen {{ display: flex; gap: 16px; margin-bottom: 32px; flex-wrap: wrap; }}
  .card {{ background: white; border-radius: 8px; padding: 16px 20px; box-shadow: 0 1px 3px rgba(0,0,0,.1); min-width: 200px; flex: 1; }}
  .card .label {{ font-size: 12px; color: #666; text-transform: uppercase; }}
  .card .valor {{ font-size: 22px; font-weight: 600; margin-top: 4px; }}
  .charts {{ display: flex; gap: 24px; flex-wrap: wrap; margin-bottom: 24px; }}
  .chart-box {{ background: white; border-radius: 8px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,.1);
                flex: 1; min-width: 320px; }}
  .chart-box.full {{ flex-basis: 100%; }}
  canvas {{ max-height: 420px; }}
</style>
</head>
<body>
  <div class="header">
    <h1>FAMM - Dashboard Financiero</h1>
    <select id="selectorAnio"></select>
  </div>
  <div class="fecha">Actualizado: {fecha_actualizacion}</div>

  <div class="resumen">
    <div class="card"><div class="label">Ingresos totales</div><div class="valor" id="valIngresos">-</div></div>
    <div class="card"><div class="label">Gastos totales</div><div class="valor" id="valGastos">-</div></div>
    <div class="card"><div class="label">Balance</div><div class="valor" id="valBalance">-</div></div>
    <div class="card"><div class="label">Ingresos: Compensaciones Ambientales</div><div class="valor" id="valCompensacion">-</div></div>
    <div class="card"><div class="label">Ingresos: Asociados y Proyectos</div><div class="valor" id="valAsociados">-</div></div>
  </div>

  <div class="charts">
    <div class="chart-box full">
      <h3>Comparativo mensual (todos los años)</h3>
      <div style="margin-bottom:8px;">
        <label><input type="radio" name="tipoMensual" value="ingresos" checked> Ingresos</label>
        <label style="margin-left:16px;"><input type="radio" name="tipoMensual" value="gastos"> Gastos</label>
      </div>
      <canvas id="mensualChart"></canvas>
    </div>
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
const DATA = {data_json};

const coloresLinea = ['#2563eb', '#c0392b', '#27ae60', '#f39c12', '#8e44ad', '#16a085'];

function fmt(n) {{
  return '$' + n.toLocaleString('es-MX', {{minimumFractionDigits: 2, maximumFractionDigits: 2}});
}}

const selector = document.getElementById('selectorAnio');
DATA.anios.forEach(a => {{
  const opt = document.createElement('option');
  opt.value = a; opt.textContent = a;
  selector.appendChild(opt);
}});

let gastosChart, ingresosChart, mensualChart;

function renderAnio(anio) {{
  const gastos = DATA.gastos_por_proyecto[anio] || {{}};
  const ingresos = DATA.ingresos_por_fuente[anio] || {{}};
  const totalGastos = Object.values(gastos).reduce((a,b) => a+b, 0);
  const totalIngresos = Object.values(ingresos).reduce((a,b) => a+b, 0);

  document.getElementById('valIngresos').textContent = fmt(totalIngresos);
  document.getElementById('valGastos').textContent = fmt(totalGastos);
  document.getElementById('valBalance').textContent = fmt(totalIngresos - totalGastos);
  document.getElementById('valCompensacion').textContent = fmt(DATA.ingresos_compensacion[anio] || 0);
  document.getElementById('valAsociados').textContent = fmt(DATA.ingresos_asociados[anio] || 0);

  if (gastosChart) gastosChart.destroy();
  gastosChart = new Chart(document.getElementById('gastosChart'), {{
    type: 'bar',
    data: {{ labels: Object.keys(gastos), datasets: [{{ label: 'Gasto (MXN)', data: Object.values(gastos), backgroundColor: '#c0392b' }}] }},
    options: {{ indexAxis: 'y', plugins: {{ legend: {{ display: false }} }} }}
  }});

  if (ingresosChart) ingresosChart.destroy();
  ingresosChart = new Chart(document.getElementById('ingresosChart'), {{
    type: 'bar',
    data: {{ labels: Object.keys(ingresos), datasets: [{{ label: 'Ingreso (MXN)', data: Object.values(ingresos), backgroundColor: '#27ae60' }}] }},
    options: {{ indexAxis: 'y', plugins: {{ legend: {{ display: false }} }} }}
  }});
}}

function renderMensual(tipo) {{
  const fuente = tipo === 'ingresos' ? DATA.monthly_ingresos : DATA.monthly_gastos;
  const datasets = DATA.anios.map((anio, i) => ({{
    label: String(anio),
    data: fuente[anio] || Array(12).fill(0),
    borderColor: coloresLinea[i % coloresLinea.length],
    backgroundColor: 'transparent',
    tension: 0.2
  }}));

  if (mensualChart) mensualChart.destroy();
  mensualChart = new Chart(document.getElementById('mensualChart'), {{
    type: 'line',
    data: {{ labels: DATA.meses_es, datasets: datasets }},
    options: {{ plugins: {{ legend: {{ display: true }} }} }}
  }});
}}

selector.addEventListener('change', () => renderAnio(selector.value));
document.querySelectorAll('input[name="tipoMensual"]').forEach(r => {{
  r.addEventListener('change', (e) => renderMensual(e.target.value));
}});

if (DATA.anios.length > 0) {{
  selector.value = DATA.anios[0];
  renderAnio(DATA.anios[0]);
  renderMensual('ingresos');
}}
</script>
</body>
</html>
"""
    return html


if __name__ == "__main__":
    main()
