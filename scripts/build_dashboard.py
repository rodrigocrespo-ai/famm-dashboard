"""
FAMM - Dashboard de Gastos e Ingresos (v4)
--------------------------------------------
Jala "Monitoreo FAMM 2026.xlsx" de Google Drive y genera dashboard.html:
  - Selector de año (tarjetas y gráficas por proyecto/categoría)
  - Comparativo mensual con checkboxes para prender/apagar cada año;
    no dibuja meses futuros del año en curso
  - Gastos por proyecto: excluye traspasos; "Otros" garantizado como
    la barra mas chica
  - Ingresos por categoría (Compensaciones / Rendimientos / cada
    proyecto por su nombre real)
  - 4 tarjetas de ingreso -- Compensaciones Ambientales, Asociados,
    Proyectos, Rendimientos -- cada una clicable: abre una vista de
    detalle con gráfica mensual de esa categoría y tabla de
    transacciones (fecha, concepto, estatus, monto -- nunca RFC,
    nombres de clientes/proveedores ni otro dato sensible)

Requiere:
  - Service account de Google con acceso de LECTURA al archivo
  - GOOGLE_SERVICE_ACCOUNT_JSON (secret en GitHub Actions)
  - SPREADSHEET_IDS (lista de IDs de archivo separados por coma -- uno por año)
"""

import os
import io
import re
import json
import datetime
import openpyxl
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SPREADSHEET_IDS = [s.strip() for s in os.environ["SPREADSHEET_IDS"].split(",") if s.strip()]

SHEET_EGRESOS = "Egreso"
SHEET_INGRESOS = "Ingreso"

TEMA_TRASPASO = "9 Traspaso entre cuentas"
TIPO_INGRESO_TRASPASO = "Traspaso entre cuentas"
PROGRAMA_TRASPASO = "Traspaso entre cuentas"
PROGRAMA_COMPENSACION = "Compensación"
PROGRAMA_RENDIMIENTOS = "Rendimientos"
PROGRAMA_ASOCIADOS = "Aportación anual"

# ---------------------------------------------------------------------
# PRESUPUESTOS APROBADOS POR PROYECTO -- Rodrigo llena esto directamente.
# Formato: { año: { "nombre EXACTO del proyecto (como aparece en PROYECTO)": monto } }
# Un proyecto sin entrada aquí simplemente no muestra comparación de
# presupuesto (solo el gasto ejercido).
# ---------------------------------------------------------------------
PRESUPUESTOS = {
    # 2026: {
    #     "2.1.1 Obras de conservación de suelos": 25000000,
    #     "2.1.2 Reforestación Compensaciones": 22000000,
    # },
}

EGRESOS_COLUMNAS = {
    "proyecto": "PROYECTO",
    "tema": "TEMA",
    "programa": "PROGRAMA",
    "concepto": "ACTIVIDAD / CONCEPTO",
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
    "cliente": "Cliente",
}

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

MESES_ES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
            "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

TARJETAS = ["compensaciones", "asociados", "proyectos", "rendimientos"]
TARJETAS_LABEL = {
    "compensaciones": "Compensaciones Ambientales",
    "asociados": "Asociados",
    "proyectos": "Proyectos",
    "rendimientos": "Rendimientos",
}

CONCEPTO_CATEGORIAS = [
    "Sueldos y salarios",
    "Insumos",
    "Implementadores",
    "Servicios profesionales",
    "Estudios técnicos",
    "Evento/Asamblea",
    "Adquisición de equipos",
    "Otros",
]

# Color fijo por categoria (paleta Economist) -- para que cada concepto
# tenga siempre el MISMO color en cualquier proyecto, no por posicion.
CONCEPTO_COLORES = {
    "Sueldos y salarios": "#e3120b",
    "Insumos": "#01295f",
    "Implementadores": "#f2909a",
    "Servicios profesionales": "#8fbfe0",
    "Estudios técnicos": "#758d99",
    "Evento/Asamblea": "#a2b1b8",
    "Adquisición de equipos": "#4d5b61",
    "Otros": "#7b3014",
}


def clasificar_concepto(texto):
    t = str(texto or "").lower()
    if "sueldo" in t or "salario" in t or "nómina" in t or "nomina" in t or "bono de productividad" in t:
        return "Sueldos y salarios"
    if any(kw in t for kw in ["biocostal", "insumo", "material", "consumible", "semilla",
                                "charola", "terrapod", "árbol", "arbol", "planta de"]):
        return "Insumos"
    if any(kw in t for kw in ["servicio", "honorario", "consultor", "asesor", "análisis de propuestas",
                                "analisis de propuestas", "evaluación", "evaluacion"]):
        return "Servicios profesionales"
    if any(kw in t for kw in ["hectárea", "hectarea", "obras de reforestación", "obras de reforestacion",
                                "obras de suelos", "plantación", "plantacion", "mantenimiento", "implementación", "implementacion"]):
        return "Implementadores"
    if re.search(r'\bha\b', t):
        return "Implementadores"
    if any(kw in t for kw in ["libro blanco", "estudio", "análisis de política", "analisis de politica",
                                "análisis de zonas", "analisis de zonas"]):
        return "Estudios técnicos"
    if any(kw in t for kw in ["evento", "asamblea", "diálogo", "dialogo", "hotel", "audiovisual",
                                "montaje", "roll up", "roll-up", "disco duro", "desplegado",
                                "publicación", "publicacion", "convocatoria", "quinta real", "producción", "produccion"]):
        return "Evento/Asamblea"
    if "equipo" in t or "adquisición" in t or "adquisicion" in t:
        return "Adquisición de equipos"
    return "Otros"


def get_drive_service():
    creds_json = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def descargar_excel(service, file_id):
    request = service.files().get_media(fileId=file_id)
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
    if valor is None:
        return None, None, None
    if isinstance(valor, (datetime.datetime, datetime.date)):
        return valor.year, valor.month, valor.strftime("%d %b %Y")
    texto = str(valor).strip()
    for fmt in ("%d %b %Y", "%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            d = datetime.datetime.strptime(texto, fmt)
            return d.year, d.month, d.strftime("%d %b %Y")
        except ValueError:
            continue
    return None, None, texto


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
        anio, mes, fecha_txt = parse_fecha(fila.get("fecha"))
        fila["anio"] = anio
        fila["mes"] = mes
        fila["fecha_txt"] = fecha_txt
        fila["monto"] = parse_monto(fila.get("total"))
        filas.append(fila)
    return filas


def agrupar_otros_mas_chico(totales):
    items = sorted(totales.items(), key=lambda x: x[1], reverse=True)
    n = len(items)
    if n == 0:
        return {}
    for k in range(1, n + 1):
        if k == n:
            principales = items
            otros_sum = 0.0
            break
        cola_sum = sum(v for _, v in items[k:])
        valor_minimo_principal = items[k - 1][1]
        if cola_sum < valor_minimo_principal:
            principales = items[:k]
            otros_sum = cola_sum
            break
    else:
        principales = items
        otros_sum = 0.0
    resultado = dict(principales)
    if otros_sum > 0:
        resultado["Otros"] = otros_sum
    return dict(sorted(resultado.items(), key=lambda x: x[1], reverse=True))


def clasificar_tarjeta(texto_combinado):
    """Recibe la union en minusculas de Programa/Proyecto + Tipo de Ingreso +
    Fuente de Ingreso -- algunos años (2022-2024) traen 'Programa/Proyecto'
    vacio o 'N/A' para filas de Rendimientos, y la palabra real solo aparece
    en Tipo/Fuente de Ingreso."""
    t = texto_combinado
    if "compensaci" in t:
        return "compensaciones"
    if "rendimiento" in t:
        return "rendimientos"
    if "aportación anual" in t or "aportacion anual" in t or "asociad" in t:
        return "asociados"
    return "proyectos"


def truncar_futuro(valores, anio, hoy):
    if anio == hoy.year:
        for idx in range(hoy.month, 12):
            valores[idx] = None
    elif anio > hoy.year:
        valores = [None] * 12
    return valores


def main():
    service = get_drive_service()

    egresos_raw = []
    ingresos_raw = []
    for file_id in SPREADSHEET_IDS:
        excel_bytes = descargar_excel(service, file_id)
        wb = openpyxl.load_workbook(excel_bytes, data_only=True, read_only=True)
        egresos_raw.extend(extraer_filtrado(wb[SHEET_EGRESOS], EGRESOS_COLUMNAS))
        ingresos_raw.extend(extraer_filtrado(wb[SHEET_INGRESOS], INGRESOS_COLUMNAS))

    # "Traspaso entre cuentas" viene con mayusculas/minusculas inconsistentes
    # segun el año (ej. 'traspaso entre cuentas', 'TRaspaso entre cuentas').
    # 2022 ademas no tiene columna TEMA -- el codigo de traspaso vive en PROYECTO.
    egresos = [f for f in egresos_raw
               if "traspaso" not in (str(f.get("tema", "")) + " " + str(f.get("proyecto", ""))).lower()]
    ingresos = [f for f in ingresos_raw
                if "traspaso" not in (str(f.get("tipo", "")) + " " + str(f.get("fuente", "")) + " "
                                       + str(f.get("programa_proyecto", ""))).lower()]

    anios = sorted({f["anio"] for f in egresos + ingresos if f["anio"]}, reverse=True)
    hoy = datetime.datetime.now()

    gastos_por_proyecto_anio = {}
    conceptos_por_proyecto_anio = {}
    ingresos_por_categoria_anio = {}
    monthly_ingresos = {}
    monthly_gastos = {}
    cards_anio = {}
    monthly_por_card = {t: {} for t in TARJETAS}
    detalle_por_card = {t: {} for t in TARJETAS}

    for anio in anios:
        egresos_anio = [f for f in egresos if f["anio"] == anio]
        ingresos_anio = [f for f in ingresos if f["anio"] == anio]

        # Gastos por proyecto
        totales_proyecto = {}
        for f in egresos_anio:
            proyecto = str(f.get("proyecto") or "Sin clasificar").strip() or "Sin clasificar"
            totales_proyecto[proyecto] = totales_proyecto.get(proyecto, 0.0) + f["monto"]
        gastos_por_proyecto_anio[anio] = agrupar_otros_mas_chico(totales_proyecto)

        # Desglose por concepto, para cada proyecto (para la grafica de pay en el detalle)
        conceptos_por_proyecto = {}
        for f in egresos_anio:
            proyecto = str(f.get("proyecto") or "Sin clasificar").strip() or "Sin clasificar"
            categoria = clasificar_concepto(f.get("concepto"))
            conceptos_por_proyecto.setdefault(proyecto, {})
            conceptos_por_proyecto[proyecto][categoria] = conceptos_por_proyecto[proyecto].get(categoria, 0.0) + f["monto"]
        conceptos_por_proyecto_anio[anio] = conceptos_por_proyecto

        meses_gastos = [0.0] * 12
        for f in egresos_anio:
            if f["mes"]:
                meses_gastos[f["mes"] - 1] += f["monto"]
        monthly_gastos[anio] = truncar_futuro(meses_gastos, anio, hoy)

        # Ingresos: descarta traspasos puros y vacios; clasifica en 4 tarjetas
        ingresos_validos = []
        for f in ingresos_anio:
            cat_raw = str(f.get("programa_proyecto") or "").strip()
            tipo_raw = str(f.get("tipo") or "").strip()
            fuente_raw = str(f.get("fuente") or "").strip()
            texto_combinado = f"{cat_raw} {tipo_raw} {fuente_raw}".lower()
            if "traspaso" in texto_combinado or cat_raw == "":
                continue
            # Si Programa/Proyecto viene "N/A", usa Tipo de Ingreso como etiqueta
            f["_categoria_raw"] = cat_raw if cat_raw.upper() != "N/A" else (tipo_raw or "Sin clasificar")
            f["_tarjeta"] = clasificar_tarjeta(texto_combinado)
            ingresos_validos.append(f)

        # Ingresos por categoria (para la grafica de barras: Compensaciones/Rendimientos/proyecto)
        totales_categoria = {}
        for f in ingresos_validos:
            label = "Compensaciones" if "compensaci" in f["_categoria_raw"].lower() else f["_categoria_raw"]
            totales_categoria[label] = totales_categoria.get(label, 0.0) + f["monto"]
        ingresos_por_categoria_anio[anio] = dict(sorted(totales_categoria.items(), key=lambda x: x[1], reverse=True))

        meses_ingresos = [0.0] * 12
        for f in ingresos_validos:
            if f["mes"]:
                meses_ingresos[f["mes"] - 1] += f["monto"]
        monthly_ingresos[anio] = truncar_futuro(meses_ingresos, anio, hoy)

        # 4 tarjetas: totales, mensual, y detalle de transacciones
        cards_anio[anio] = {}
        for tarjeta in TARJETAS:
            filas_tarjeta = [f for f in ingresos_validos if f["_tarjeta"] == tarjeta]
            cards_anio[anio][tarjeta] = sum(f["monto"] for f in filas_tarjeta)

            meses_tarjeta = [0.0] * 12
            for f in filas_tarjeta:
                if f["mes"]:
                    meses_tarjeta[f["mes"] - 1] += f["monto"]
            monthly_por_card[tarjeta][anio] = truncar_futuro(meses_tarjeta, anio, hoy)

            detalle = [
                {
                    "fecha": f["fecha_txt"],
                    "donante": str(f.get("cliente") or "").strip(),
                    "monto": f["monto"],
                }
                for f in sorted(filas_tarjeta, key=lambda x: (x["mes"] or 0))
            ]
            detalle_por_card[tarjeta][anio] = detalle

    fecha_actualizacion = hoy.strftime("%d %b %Y, %H:%M")

    data_js = {
        "anios": anios,
        "gastos_por_proyecto": gastos_por_proyecto_anio,
        "conceptos_por_proyecto": conceptos_por_proyecto_anio,
        "ingresos_por_categoria": ingresos_por_categoria_anio,
        "cards": cards_anio,
        "cards_labels": TARJETAS_LABEL,
        "monthly_ingresos": monthly_ingresos,
        "monthly_gastos": monthly_gastos,
        "monthly_por_card": monthly_por_card,
        "detalle_por_card": detalle_por_card,
        "meses_es": MESES_ES,
        "concepto_colores": CONCEPTO_COLORES,
        "presupuestos": PRESUPUESTOS,
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
<script src="https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-annotation/3.0.1/chartjs-plugin-annotation.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-datalabels/2.2.0/chartjs-plugin-datalabels.min.js"></script>
<style>
  body {{ font-family: -apple-system, Segoe UI, Arial, sans-serif; margin: 0; padding: 24px;
         background: #f5f6f7; color: #1a1a1a; }}
  h1 {{ font-size: 20px; margin-bottom: 4px; display: inline-block; }}
  .header {{ display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; }}
  .fecha {{ color: #666; font-size: 13px; margin-bottom: 24px; }}
  select {{ font-size: 15px; padding: 6px 10px; border-radius: 6px; border: 1px solid #ccc; background: white; }}
  .resumen {{ display: flex; gap: 16px; margin-bottom: 32px; flex-wrap: wrap; }}
  .card {{ background: white; border-radius: 8px; padding: 16px 20px; box-shadow: 0 1px 3px rgba(0,0,0,.1);
           min-width: 180px; flex: 1; }}
  .card.clicable {{ cursor: pointer; transition: box-shadow .15s; }}
  .card.clicable:hover {{ box-shadow: 0 2px 8px rgba(0,0,0,.18); }}
  .card .label {{ font-size: 12px; color: #666; text-transform: uppercase; }}
  .card .valor {{ font-size: 22px; font-weight: 600; margin-top: 4px; }}
  .charts {{ display: flex; gap: 24px; flex-wrap: wrap; margin-bottom: 24px; }}
  .chart-box {{ background: white; border-radius: 8px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,.1);
                flex: 1; min-width: 320px; }}
  .chart-box.full {{ flex-basis: 100%; }}
  #chartMensualBox {{ border-top: 4px solid #e3120b; }}
  canvas {{ max-height: 420px; }}
  .anios-check {{ display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 8px; }}
  .anios-check label {{ font-size: 14px; }}
  table {{ width: auto; min-width: 320px; border-collapse: collapse; margin-top: 16px; font-size: 15px; }}
  th, td {{ text-align: left; padding: 6px 24px 6px 0; border-bottom: 1px solid #eee; }}
  th:last-child, td:last-child {{ padding-right: 0; }}
  th {{ color: #666; font-size: 12px; text-transform: uppercase; }}
  .volver {{ background: none; border: none; color: #2563eb; font-size: 14px; cursor: pointer; padding: 0;
             margin-bottom: 16px; }}
</style>
</head>
<body>

<div id="vistaPrincipal">
  <div class="header">
    <h1>FAMM - Dashboard Financiero</h1>
    <select id="selectorAnio"></select>
  </div>
  <div class="fecha">Actualizado: {fecha_actualizacion}</div>

  <div class="resumen">
    <div class="card"><div class="label">Ingresos totales</div><div class="valor" id="valIngresos">-</div></div>
    <div class="card"><div class="label">Gastos totales</div><div class="valor" id="valGastos">-</div></div>
    <div class="card"><div class="label">Balance</div><div class="valor" id="valBalance">-</div></div>
  </div>

  <div class="resumen" id="tarjetasIngreso"></div>

  <div class="charts">
    <div class="chart-box full" id="chartMensualBox">
      <h3>Comparativo mensual por año</h3>
      <div class="anios-check" id="aniosCheck"></div>
      <div style="margin-bottom:8px;">
        <label><input type="radio" name="tipoMensual" value="ingresos" checked> Ingresos</label>
        <label style="margin-left:16px;"><input type="radio" name="tipoMensual" value="gastos"> Gastos</label>
      </div>
      <canvas id="mensualChart"></canvas>
      <div id="totalesMensual" style="margin-top:8px;"></div>
    </div>
  </div>

  <div class="charts">
    <div class="chart-box">
      <h3>Ingresos por categoría</h3>
      <div class="anios-check" id="aniosCheckCat"></div>
      <canvas id="ingresosChart"></canvas>
    </div>
    <div class="chart-box" style="min-width: 300px;">
      <h3>Ingresos de compensaciones por mes</h3>
      <div class="anios-check" id="aniosCheckComp"></div>
      <canvas id="mensualCompensacionesChart"></canvas>
      <div id="totalesComp" style="margin-top:8px;"></div>
    </div>
  </div>

  <div class="charts">
    <div class="chart-box" style="min-width: 300px;">
      <h3>Ingresos por cuotas de socios</h3>
      <div class="anios-check" id="aniosCheckAsoc"></div>
      <canvas id="mensualAsociadosChart"></canvas>
      <div id="totalesAsoc" style="margin-top:8px;"></div>
    </div>
    <div class="chart-box" style="min-width: 300px;">
      <h3>Ingresos por proyectos</h3>
      <div class="anios-check" id="aniosCheckProy"></div>
      <canvas id="mensualProyectosChart"></canvas>
      <div id="totalesProy" style="margin-top:8px;"></div>
    </div>
  </div>

  <div class="charts">
    <div class="chart-box full">
      <h3>Gastos por proyecto</h3>
      <canvas id="gastosChart"></canvas>
    </div>
  </div>
</div>

<div id="vistaDetalle" style="display:none;">
  <button class="volver" id="btnVolver">&larr; Volver al dashboard</button>
  <h2 id="detalleTitulo"></h2>
  <div class="chart-box full" style="margin-bottom:16px;">
    <canvas id="detalleChart"></canvas>
  </div>
  <table id="detalleTabla">
    <thead id="detalleTablaHead"></thead>
    <tbody id="detalleTablaBody"></tbody>
  </table>
</div>

<div id="vistaDetalleGasto" style="display:none;">
  <button class="volver" id="btnVolverGasto">&larr; Volver al dashboard</button>
  <h2 id="detalleGastoTitulo"></h2>
  <div class="charts">
    <div class="chart-box" style="flex: 3; min-width: 320px;">
      <canvas id="detalleGastoChart" style="max-height: 160px;"></canvas>
    </div>
    <div class="card" style="flex: 1; min-width: 180px;">
      <div class="label">% del presupuesto ejercido</div>
      <div class="valor" id="detalleGastoPct">-</div>
    </div>
  </div>
  <p id="detalleGastoResumen" style="font-size:15px;"></p>
  <div class="chart-box full" style="margin-top:16px;">
    <h3>Gasto por concepto</h3>
    <canvas id="detalleGastoPie" style="max-height: 320px;"></canvas>
  </div>
</div>

<script>
const DATA = {data_json};

Chart.defaults.font.family = "-apple-system, Segoe UI, Arial, sans-serif";
Chart.defaults.color = '#1a1a1a';
Chart.defaults.borderColor = '#e3e3e3';
Chart.defaults.plugins.legend.labels.usePointStyle = false;
Chart.defaults.plugins.legend.labels.boxWidth = 14;
Chart.register(ChartDataLabels);
Chart.defaults.set('plugins.datalabels', {{ display: false }});  // solo se activa donde se pida explicitamente


const ECONOMIST_ROJO = '#e3120b';
const ECONOMIST_AZUL = '#006ba2';
const coloresLinea = ['#e3120b', '#01295f', '#f2909a', '#8fbfe0', '#758d99', '#a2b1b8', '#4d5b61', '#7b3014'];

function fmt(n) {{
  return '$' + n.toLocaleString('es-MX', {{minimumFractionDigits: 2, maximumFractionDigits: 2}});
}}

const selector = document.getElementById('selectorAnio');
DATA.anios.forEach(a => {{
  const opt = document.createElement('option');
  opt.value = a; opt.textContent = a;
  selector.appendChild(opt);
}});

function crearCheckboxesAnio(containerId, claseCss, onChange) {{
  const cont = document.getElementById(containerId);
  DATA.anios.forEach((a, i) => {{
    const label = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox'; cb.checked = true; cb.value = a; cb.className = claseCss;
    cb.style.accentColor = coloresLinea[i % coloresLinea.length];
    cb.addEventListener('change', onChange);
    label.appendChild(cb);
    label.appendChild(document.createTextNode(' ' + a));
    cont.appendChild(label);
  }});
}}

function aniosActivosDe(claseCss) {{
  return Array.from(document.querySelectorAll('.' + claseCss + ':checked')).map(cb => cb.value);
}}

function renderTotalesAnio(containerId, fuentePorAnio, claseCss) {{
  const anioActualStr = String(new Date().getFullYear());
  const aniosActivos = aniosActivosDe(claseCss).sort((a, b) => Number(b) - Number(a));
  const cont = document.getElementById(containerId);
  cont.innerHTML = '';
  aniosActivos.forEach(anio => {{
    const valores = fuentePorAnio[anio] || [];
    const total = valores.reduce((a, b) => a + (b || 0), 0);
    const etiqueta = (anio === anioActualStr) ? ('Total acumulado ' + anio) : ('Total ' + anio);
    const i = DATA.anios.map(String).indexOf(anio);
    const color = coloresLinea[i % coloresLinea.length];
    const fila = document.createElement('div');
    fila.style.cssText = 'display:flex; justify-content:space-between; max-width:320px; font-size:16px; font-weight:600; margin-top:4px; color:' + color + ';';
    const spanLabel = document.createElement('span');
    spanLabel.textContent = etiqueta + ':';
    const spanValor = document.createElement('span');
    spanValor.textContent = fmt(total);
    fila.appendChild(spanLabel);
    fila.appendChild(spanValor);
    cont.appendChild(fila);
  }});
}}

crearCheckboxesAnio('aniosCheck', 'anioCheck', () => renderMensual(tipoMensualActual));
crearCheckboxesAnio('aniosCheckCat', 'anioCheckCat', () => renderIngresosCategoria());
crearCheckboxesAnio('aniosCheckComp', 'anioCheckComp', () => renderTarjetaCompensaciones());
crearCheckboxesAnio('aniosCheckAsoc', 'anioCheckAsoc', () => renderTarjetaAsociados());
crearCheckboxesAnio('aniosCheckProy', 'anioCheckProy', () => renderTarjetaProyectos());

let gastosChart, ingresosChart, mensualChart, detalleChart, detalleGastoChart, detalleGastoPieChart;
let mensualCompensacionesChart, mensualAsociadosChart, mensualProyectosChart;
let tipoMensualActual = 'ingresos';

function renderAnio(anio) {{
  const gastos = DATA.gastos_por_proyecto[anio] || {{}};
  const ingresos = DATA.ingresos_por_categoria[anio] || {{}};
  const totalGastos = Object.values(gastos).reduce((a,b) => a+b, 0);
  const totalIngresos = Object.values(ingresos).reduce((a,b) => a+b, 0);

  document.getElementById('valIngresos').textContent = fmt(totalIngresos);
  document.getElementById('valGastos').textContent = fmt(totalGastos);
  document.getElementById('valBalance').textContent = fmt(totalIngresos - totalGastos);

  const cont = document.getElementById('tarjetasIngreso');
  cont.innerHTML = '';
  const cardsAnio = DATA.cards[anio] || {{}};
  Object.keys(DATA.cards_labels).forEach(key => {{
    const div = document.createElement('div');
    div.className = 'card clicable';
    div.innerHTML = `<div class="label">Ingresos: ${{DATA.cards_labels[key]}}</div><div class="valor">${{fmt(cardsAnio[key] || 0)}}</div>`;
    div.addEventListener('click', () => mostrarDetalle(key, anio));
    cont.appendChild(div);
  }});

  if (gastosChart) gastosChart.destroy();
  gastosChart = new Chart(document.getElementById('gastosChart'), {{
    type: 'bar',
    data: {{ labels: Object.keys(gastos), datasets: [{{ label: 'Gasto (MXN)', data: Object.values(gastos), backgroundColor: ECONOMIST_ROJO }}] }},
    options: {{ indexAxis: 'y', plugins: {{ legend: {{ display: false }} }},
      onClick: (evt, elements) => {{
        if (elements.length > 0) {{
          const proyecto = Object.keys(gastos)[elements[0].index];
          mostrarDetalleGasto(proyecto, anio);
        }}
      }}
    }}
  }});
  document.getElementById('gastosChart').style.cursor = 'pointer';
}}

function construirDatasetsMultiAnio(fuentePorAnio, claseCss) {{
  const aniosActivos = aniosActivosDe(claseCss);
  return DATA.anios
    .filter(anio => aniosActivos.includes(String(anio)))
    .map((anio) => {{
      const i = DATA.anios.indexOf(anio);
      return {{
        label: String(anio),
        data: fuentePorAnio[anio] || Array(12).fill(null),
        backgroundColor: coloresLinea[i % coloresLinea.length]
      }};
    }});
}}

function renderMensual(tipo) {{
  tipoMensualActual = tipo;
  const fuente = tipo === 'ingresos' ? DATA.monthly_ingresos : DATA.monthly_gastos;
  const datasets = construirDatasetsMultiAnio(fuente, 'anioCheck');

  if (mensualChart) mensualChart.destroy();
  mensualChart = new Chart(document.getElementById('mensualChart'), {{
    type: 'bar',
    data: {{ labels: DATA.meses_es, datasets: datasets }},
    options: {{ plugins: {{ legend: {{ display: true }} }} }}
  }});
  renderTotalesAnio('totalesMensual', fuente, 'anioCheck');
}}

function renderIngresosCategoria() {{
  const aniosActivos = aniosActivosDe('anioCheckCat');
  const porAnio = DATA.ingresos_por_categoria;

  const totalPorCategoria = {{}};
  aniosActivos.forEach(anio => {{
    const datos = porAnio[anio] || {{}};
    Object.keys(datos).forEach(cat => {{
      totalPorCategoria[cat] = (totalPorCategoria[cat] || 0) + datos[cat];
    }});
  }});

  const categorias = Object.keys(totalPorCategoria).sort((a, b) => totalPorCategoria[b] - totalPorCategoria[a]);
  const TOP_N = 12;
  let categoriasFinal = categorias;
  let categoriasOtros = [];
  if (categorias.length > TOP_N) {{
    categoriasFinal = categorias.slice(0, TOP_N);
    categoriasOtros = categorias.slice(TOP_N);
    categoriasFinal = categoriasFinal.concat(['Otros']);
  }}

  const datasets = aniosActivos.map(anio => {{
    const i = DATA.anios.map(String).indexOf(anio);
    const datos = porAnio[anio] || {{}};
    const valores = categoriasFinal.map(cat => {{
      if (cat === 'Otros') {{
        return categoriasOtros.reduce((suma, c) => suma + (datos[c] || 0), 0);
      }}
      return datos[cat] || 0;
    }});
    return {{ label: String(anio), data: valores, backgroundColor: coloresLinea[i % coloresLinea.length] }};
  }});

  if (ingresosChart) ingresosChart.destroy();
  ingresosChart = new Chart(document.getElementById('ingresosChart'), {{
    type: 'bar',
    data: {{ labels: categoriasFinal, datasets: datasets }},
    options: {{ indexAxis: 'y', plugins: {{ legend: {{ display: true }} }} }}
  }});
}}

function renderTarjetaCompensaciones() {{
  if (mensualCompensacionesChart) mensualCompensacionesChart.destroy();
  mensualCompensacionesChart = new Chart(document.getElementById('mensualCompensacionesChart'), {{
    type: 'bar',
    data: {{ labels: DATA.meses_es, datasets: construirDatasetsMultiAnio(DATA.monthly_por_card.compensaciones || {{}}, 'anioCheckComp') }},
    options: {{ plugins: {{ legend: {{ display: true }} }} }}
  }});
  renderTotalesAnio('totalesComp', DATA.monthly_por_card.compensaciones || {{}}, 'anioCheckComp');
}}

function renderTarjetaAsociados() {{
  if (mensualAsociadosChart) mensualAsociadosChart.destroy();
  mensualAsociadosChart = new Chart(document.getElementById('mensualAsociadosChart'), {{
    type: 'bar',
    data: {{ labels: DATA.meses_es, datasets: construirDatasetsMultiAnio(DATA.monthly_por_card.asociados || {{}}, 'anioCheckAsoc') }},
    options: {{ plugins: {{ legend: {{ display: true }} }} }}
  }});
  renderTotalesAnio('totalesAsoc', DATA.monthly_por_card.asociados || {{}}, 'anioCheckAsoc');
}}

function renderTarjetaProyectos() {{
  if (mensualProyectosChart) mensualProyectosChart.destroy();
  mensualProyectosChart = new Chart(document.getElementById('mensualProyectosChart'), {{
    type: 'bar',
    data: {{ labels: DATA.meses_es, datasets: construirDatasetsMultiAnio(DATA.monthly_por_card.proyectos || {{}}, 'anioCheckProy') }},
    options: {{ plugins: {{ legend: {{ display: true }} }} }}
  }});
  renderTotalesAnio('totalesProy', DATA.monthly_por_card.proyectos || {{}}, 'anioCheckProy');
}}

function mostrarDetalle(tarjeta, anio) {{
  document.getElementById('vistaPrincipal').style.display = 'none';
  document.getElementById('vistaDetalle').style.display = 'block';
  document.getElementById('detalleTitulo').textContent = DATA.cards_labels[tarjeta] + ' - ' + anio;

  const mensual = (DATA.monthly_por_card[tarjeta] && DATA.monthly_por_card[tarjeta][anio]) || Array(12).fill(null);
  if (detalleChart) detalleChart.destroy();
  detalleChart = new Chart(document.getElementById('detalleChart'), {{
    type: 'line',
    data: {{ labels: DATA.meses_es, datasets: [{{ label: DATA.cards_labels[tarjeta], data: mensual,
             borderColor: ECONOMIST_AZUL, backgroundColor: 'rgba(0,107,162,.12)', fill: true, spanGaps: false, tension: 0.2 }}] }},
    options: {{ plugins: {{ legend: {{ display: false }} }} }}
  }});

  const detalle = (DATA.detalle_por_card[tarjeta] && DATA.detalle_por_card[tarjeta][anio]) || [];
  const conDonante = tarjeta !== 'rendimientos';
  const thead = document.getElementById('detalleTablaHead');
  thead.innerHTML = conDonante
    ? '<tr><th>Fecha</th><th>Donante</th><th>Monto</th></tr>'
    : '<tr><th>Fecha</th><th>Monto</th></tr>';

  const tbody = document.getElementById('detalleTablaBody');
  tbody.innerHTML = '';
  detalle.forEach(row => {{
    const tr = document.createElement('tr');
    tr.innerHTML = conDonante
      ? `<td>${{row.fecha || ''}}</td><td>${{row.donante || ''}}</td><td>${{fmt(row.monto)}}</td>`
      : `<td>${{row.fecha || ''}}</td><td>${{fmt(row.monto)}}</td>`;
    tbody.appendChild(tr);
  }});
}}

document.getElementById('btnVolver').addEventListener('click', () => {{
  document.getElementById('vistaDetalle').style.display = 'none';
  document.getElementById('vistaPrincipal').style.display = 'block';
}});

function mostrarDetalleGasto(proyecto, anio) {{
  if (proyecto === 'Otros') {{
    alert('"Otros" agrupa varios proyectos pequeños -- no tiene un presupuesto individual. Selecciona un proyecto especifico para ver su comparacion.');
    return;
  }}
  document.getElementById('vistaPrincipal').style.display = 'none';
  document.getElementById('vistaDetalleGasto').style.display = 'block';
  document.getElementById('detalleGastoTitulo').textContent = proyecto + ' - ' + anio;

  const ejercido = (DATA.gastos_por_proyecto[anio] || {{}})[proyecto] || 0;
  const presupuestosAnio = DATA.presupuestos[anio] || {{}};
  const presupuesto = presupuestosAnio[proyecto];
  const hayPresupuesto = presupuesto !== undefined && presupuesto !== null;

  if (detalleGastoChart) detalleGastoChart.destroy();
  const labels = hayPresupuesto ? ['Gasto ejercido', 'Presupuesto aprobado'] : ['Gasto ejercido'];
  const data = hayPresupuesto ? [ejercido, presupuesto] : [ejercido];
  const colores = hayPresupuesto ? [ECONOMIST_ROJO, ECONOMIST_AZUL] : [ECONOMIST_ROJO];

  const anotaciones = {{}};
  if (hayPresupuesto) {{
    anotaciones.lineaPresupuesto = {{
      type: 'line',
      xMin: presupuesto,
      xMax: presupuesto,
      borderColor: '#1a1a1a',
      borderWidth: 2,
      borderDash: [6, 4],
      label: {{ display: true, content: 'Presupuesto máx.', position: 'end',
                backgroundColor: '#1a1a1a', color: '#fff', font: {{ size: 10 }} }}
    }};
  }}

  detalleGastoChart = new Chart(document.getElementById('detalleGastoChart'), {{
    type: 'bar',
    data: {{ labels: labels, datasets: [{{ data: data, backgroundColor: colores, barThickness: 22 }}] }},
    options: {{ indexAxis: 'y', plugins: {{ legend: {{ display: false }}, annotation: {{ annotations: anotaciones }} }} }}
  }});

  const pctBox = document.getElementById('detalleGastoPct');
  const resumen = document.getElementById('detalleGastoResumen');
  if (hayPresupuesto && presupuesto > 0) {{
    const pct = (ejercido / presupuesto * 100).toFixed(1);
    pctBox.textContent = pct + '%';
    resumen.textContent = 'Ejercido: ' + fmt(ejercido) + ' de ' + fmt(presupuesto) + ' presupuestados.';
  }} else {{
    pctBox.textContent = 'N/D';
    resumen.textContent = 'Gasto ejercido: ' + fmt(ejercido) + '. Presupuesto aprobado: aun no capturado para este proyecto.';
  }}

  const conceptos = (DATA.conceptos_por_proyecto[anio] && DATA.conceptos_por_proyecto[anio][proyecto]) || {{}};
  const coloresPorConcepto = Object.keys(conceptos).map(k => DATA.concepto_colores[k] || '#a2b1b8');
  if (detalleGastoPieChart) detalleGastoPieChart.destroy();
  detalleGastoPieChart = new Chart(document.getElementById('detalleGastoPie'), {{
    type: 'pie',
    data: {{
      labels: Object.keys(conceptos),
      datasets: [{{ data: Object.values(conceptos), backgroundColor: coloresPorConcepto }}]
    }},
    options: {{
      plugins: {{
        legend: {{ display: true, position: 'right' }},
        tooltip: {{
          callbacks: {{
            label: (ctx) => {{
              const total = ctx.dataset.data.reduce((a,b) => a+b, 0);
              const pct = (ctx.parsed / total * 100).toFixed(1);
              return ctx.label + ': ' + fmt(ctx.parsed) + ' (' + pct + '%)';
            }}
          }}
        }},
        datalabels: {{
          display: (ctx) => {{
            const total = ctx.dataset.data.reduce((a,b) => a+b, 0);
            const pct = total ? (ctx.dataset.data[ctx.dataIndex] / total * 100) : 0;
            return pct >= 5;
          }},
          color: '#fff',
          textStrokeColor: '#000',
          textStrokeWidth: 2,
          font: {{ weight: 'bold', size: 13 }},
          formatter: (value, ctx) => {{
            const total = ctx.dataset.data.reduce((a,b) => a+b, 0);
            const pct = total ? (value / total * 100) : 0;
            return pct.toFixed(0) + '%';
          }}
        }}
      }}
    }}
  }});
}}

document.getElementById('btnVolverGasto').addEventListener('click', () => {{
  document.getElementById('vistaDetalleGasto').style.display = 'none';
  document.getElementById('vistaPrincipal').style.display = 'block';
}});

selector.addEventListener('change', () => renderAnio(selector.value));
document.querySelectorAll('input[name="tipoMensual"]').forEach(r => {{
  r.addEventListener('change', (e) => renderMensual(e.target.value));
}});

if (DATA.anios.length > 0) {{
  selector.value = DATA.anios[0];
  renderAnio(DATA.anios[0]);
  renderMensual('ingresos');
  renderIngresosCategoria();
  renderTarjetaCompensaciones();
  renderTarjetaAsociados();
  renderTarjetaProyectos();
}}
</script>
</body>
</html>
"""
    return html


if __name__ == "__main__":
    main()
