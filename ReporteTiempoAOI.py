
import io
from datetime import datetime, date, time, timedelta
from pathlib import Path
from typing import Callable, Dict, Tuple, Optional
from io import BytesIO
import base64
import requests



import pandas as pd
import streamlit as st

import math

import calendar

# Función auxiliar (ponla al inicio, después de los imports)
def mes_en_espanol(fecha: datetime) -> str:
    # Obtiene el número del mes (1 = enero, 12 = diciembre)
    mes_num = fecha.month
    # Usa calendar para obtener nombre en inglés, luego traduce
    mes_ing = fecha.strftime("%B")
    meses_es = [
        "Enero", "Febrero", "Marzo", "Abril",
        "Mayo", "Junio", "Julio", "Agosto",
        "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]
    return meses_es[mes_num - 1]



st.set_page_config(page_title="Reporte Novedades de Tiempo", layout="wide")
pd.options.display.float_format = "{:.2f}".format

DIAS_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
DAY_06 = time(6, 0)
DAY_18 = time(18, 0)
LUNCH_I, LUNCH_F = time(12, 0), time(13, 0)
DEFAULT_BASE_FESTIVA_HORAS = 8

MinutesDict = Dict[str, int]

def _to_time(x):
    if pd.isna(x) or x is None:
        return None
    if isinstance(x, time):
        return x
    try:
        dt = pd.to_datetime(str(x), dayfirst=False, errors="coerce")
        return None if pd.isna(dt) else dt.time()
    except Exception:
        return None

def week_dates(year: int, week: int):
    monday = datetime.fromisocalendar(year, week, 1)
    return [monday + timedelta(days=i) for i in range(7)]

def minutes_between(a: datetime, b: datetime) -> int:
    return max(0, int((b - a).total_seconds() // 60))

def clamp_interval(a0: datetime, a1: datetime, b0: datetime, b1: datetime):
    s = max(a0, b0)
    e = min(a1, b1)
    return (s, e) if s < e else (None, None)

@st.cache_data(show_spinner=False)
def load_tables(xlsx):
    xl = pd.ExcelFile(xlsx)
    def get(name):
        return xl.parse(name)

    personal = get("Personal")
    j_diaria = get("JornadaDiaria")
    j_horaria = get("JornadaHoraria")

    for col in ["HoraInicioLabor", "HoraFinLabor", "HoraInicioAlmuerzo", "HoraFinAlmuerzo"]:
        if col in j_horaria.columns:
            j_horaria[col] = j_horaria[col].apply(_to_time)

    if "Dia" in j_horaria.columns:
        j_horaria["Dia"] = j_horaria["Dia"].astype(str).str.strip()
    if "Grupo" in j_horaria.columns:
        j_horaria["Grupo"] = j_horaria["Grupo"].astype(str).str.strip()

    festivos = get("Festivos") if "Festivos" in xl.sheet_names else pd.DataFrame(columns=["Fecha", "Descripción"])
    config = get("Config") if "Config" in xl.sheet_names else pd.DataFrame({
        "Parametro": ["Año", "Semana", "Tarifa Subsidio (por día)"],
        "Valor": [datetime.now().year, 1, 0]
    })
    return personal, j_diaria, j_horaria, festivos, config

def dia_es(d: date):
    return DIAS_ES[d.weekday()]

# =========================================
# Marca de agua EPM visible sin superponerse
# =========================================
LOGO_EPM_URL = "https://upload.wikimedia.org/wikipedia/commons/thumb/6/62/Logo_EPM.svg/990px-Logo_EPM.svg.png"

def _inject_epm_watermark():
    st.markdown(f"""
        <style>
        /* 1) Fondo de .stApp = velo translúcido + logo EPM centrado y fijo */
        .stApp {{
            /* Capa 1: velo blanco translúcido para atenuar el logo (ajusta 0.90–0.98) */
            background-image:
                linear-gradient(rgba(255,255,255,0.88), rgba(255,255,255,0.70)),
                url('{LOGO_EPM_URL}');
            background-repeat: no-repeat, no-repeat;
            background-position: center 45%, center 50%;
            background-size: cover, 30%;
            background-attachment: fixed, fixed; /* estable al scroll */
        }}

        /* 2) Volver transparentes contenedores que traen fondo por tema */
        .stApp .main, 
        .stApp .block-container {{
            background: transparent !important;
        }}
        /* Sidebar / header suelen tener su propio fondo; los dejamos como están */
        </style>
    """, unsafe_allow_html=True)

_inject_epm_watermark()


st.title("Reporte de Novedades de Tiempo — Área Operaciones Ituango EPM 📅")
st.caption("Clasificación automática: OD/ON/ED/EN/FD/FN/EFD/EFN + ausentismo y subsidio.")


with st.sidebar:
    st.header("Configuración")
    default_path = Path("TablasReporteTiempoOperaciones.xlsx")
    up = st.file_uploader("Sube 'TablasReporteTiempoOperaciones.xlsx'", type=["xlsx"])
    fuente = up if up else (default_path if default_path.exists() else None)
    if not fuente:
        st.warning("Falta el maestro de tablas. Sube el archivo con ese nombre.")
        st.stop()

    personal, j_diaria, j_horaria, festivos_df, config = load_tables(fuente)

    col_nombre = None
    for cand in ["Nombre", "Funcionario", "Nombre Funcionario", "NOMBRE"]:
        if cand in personal.columns:
            col_nombre = cand
            break
    if col_nombre is None:
        st.error("La hoja 'Personal' debe tener una columna de nombre (e.g., 'Nombre').")
        st.stop()

    year_default = int(config.loc[config["Parametro"] == "Año", "Valor"].values[0]) if "Parametro" in config else datetime.now().year
    week_default = int(config.loc[config["Parametro"] == "Semana", "Valor"].values[0]) if "Parametro" in config else datetime.now().isocalendar()[1]

    year = st.number_input("Año (ISO)", min_value=2000, max_value=2100, value=year_default, step=1)
    week = st.number_input("Semana ISO", min_value=1, max_value=53, value=week_default, step=1)

    tarifa_subsidio = 0.0
    if "Parametro" in config and any(config["Parametro"].str.contains("Subsidio", case=False)):
        tarifa_subsidio = float(config.loc[config["Parametro"].str.contains("Subsidio", case=False), "Valor"].values[0])
    tarifa_subsidio = st.number_input("Tarifa Subsidio (por día)", min_value=0.0, value=float(tarifa_subsidio), step=1.0)
    st.session_state["tarifa_subsidio"] = tarifa_subsidio

    descontar_almuerzo_default = st.toggle("Descontar almuerzo típico (12:00–13:00)", value=True)
    validar_semana = st.toggle("Validar que el registro esté dentro de la semana", value=True)

if "nov_df" not in st.session_state:
    st.session_state.nov_df = pd.DataFrame(columns=[
        "Nombre del Funcionario", "Especialidad", "Grupo", "Equipo",
        "Mes", "Día",
        "Fecha y hora de inicio labores", "Fecha y hora final labores",
        "Actividad / Frente de trabajo",
        "Subsidio de Transporte (días)",
        "Ausentismo - Concepto", "Ausentismo - Hora Inicio", "Ausentismo - Hora Fin", "Ausentismo - Justificación",
        "Descontar Almuerzo"
    ])

st.subheader("1) Agregar registro")
nombres = sorted(personal[col_nombre].dropna().astype(str).unique().tolist())
opciones_nombre = ["(Escribir manualmente)"] + nombres
col1, col2 = st.columns([2, 1])
with col1:
    nombre_sel = st.selectbox("Nombre del Funcionario (o elige 'Escribir manualmente')", opciones_nombre, index=0)
with col2:
    filtro = st.text_input("Buscar/Escribir nombre", value="")

sugerencias = []
if filtro:
    fl = filtro.strip().lower()
    sugerencias = [n for n in nombres if fl in n.lower()][:10]

nombre_final = nombre_sel if nombre_sel != "(Escribir manualmente)" else (filtro if filtro else None)

esp_val = grp_val = eq_val = ""
if nombre_final:
    pers_idx = personal[personal[col_nombre].astype(str).str.lower() == str(nombre_final).lower()]
    if not pers_idx.empty:
        esp_val = pers_idx.iloc[0]["Especialidad"] if "Especialidad" in pers_idx.columns else ""
        grp_val = pers_idx.iloc[0]["Grupo"] if "Grupo" in pers_idx.columns else ""
        eq_val = pers_idx.iloc[0]["Equipo"] if "Equipo" in pers_idx.columns else ""

cA, cB, cC = st.columns(3)
with cA:
    esp_in = st.text_input("Especialidad", value=str(esp_val))
with cB:
    grp_in = st.text_input("Grupo", value=str(grp_val))
with cC:
    eq_in = st.text_input("Equipo", value=str(eq_val))

cD, cE = st.columns(2)
with cD:
    fechain = st.date_input("Fecha de inicio de labores")
    horain = st.time_input("Hora de inicio de labores")
    fecha_ini = datetime.combine(fechain, horain)
with cE:
    fechafin = st.date_input("Fecha final de labores")
    horafin = st.time_input("Hora final de labores")
    fecha_fin = datetime.combine(fechafin, horafin)

act = st.text_input("Actividad / Frente de trabajo", value="")
sub_dias = st.number_input("Subsidio de Transporte (días)", min_value=0, step=1, value=0)

descontar_almuerzo_reg = st.checkbox("Descontar almuerzo en este registro", value=descontar_almuerzo_default)

st.markdown("**Ausentismo**")
a1, a2, a3, a4 = st.columns([1, 1, 1, 2])
with a1:
    aus_c = st.text_input("Concepto", value="")
with a2:
    aus_i = None
    if st.checkbox("Ingresar fecha/hora de inicio"):
        ausfechaini = st.date_input("Fecha de inicio", key="ausfechaini")
        aushoraini = st.time_input("Hora de inicio", value=time(9, 0), key="aushoraini")
        aus_i = datetime.combine(ausfechaini, aushoraini)
with a3:
    aus_f = None
    if st.checkbox("Ingresar fecha/hora de fin"):
        ausfechafin = st.date_input("Fecha de fin", key="ausfechafin")
        aushorafin = st.time_input("Hora de fin", value=time(9, 0), key="aushorafin")
        aus_f = datetime.combine(ausfechafin, aushorafin)
with a4:
    aus_j = st.text_input("Justificación", value="")

if st.button("➕ Agregar"):
    if not nombre_final:
        st.error("Debes seleccionar o escribir el nombre del funcionario.")
    elif fecha_fin <= fecha_ini:
        st.error("La fecha/hora final debe ser mayor que la inicial.")
    else:
        d = fecha_ini.date()
        fila = {
            "Nombre del Funcionario": nombre_final,
            "Especialidad": esp_in,
            "Grupo": grp_in,
            "Equipo": eq_in,
            "Mes": mes_en_espanol(fecha_ini),
            "Día": DIAS_ES[d.weekday()],
            "Fecha y hora de inicio labores": fecha_ini,
            "Fecha y hora final labores": fecha_fin,
            "Actividad / Frente de trabajo": act,
            "Subsidio de Transporte (días)": sub_dias,
            "Ausentismo - Concepto": aus_c,
            "Ausentismo - Hora Inicio": aus_i,
            "Ausentismo - Hora Fin": aus_f,
            "Ausentismo - Justificación": aus_j,
            "Descontar Almuerzo": bool(descontar_almuerzo_reg),
        }
        st.session_state.nov_df = pd.concat([st.session_state.nov_df, pd.DataFrame([fila])], ignore_index=True)
        st.success("Registro agregado.")

if st.button("🧹 Limpiar todo"):
    st.session_state.nov_df = st.session_state.nov_df.iloc[0:0]
    st.info("Se vaciaron los registros.")

st.subheader("2) Registros capturados (editable)")
st.caption("Edita celdas directamente.")

st.session_state.nov_df = st.data_editor(
    st.session_state.nov_df,
    num_rows="dynamic",
    use_container_width=True,
    key="editor_nov",
    column_config={
        "Descontar Almuerzo": st.column_config.CheckboxColumn("Descontar Almuerzo", default=True)
    },
)

def es_festivo_fn_local(d: date) -> bool:
    if d.weekday() == 6:
        return True
    try:
        return any(pd.to_datetime(festivos_df["Fecha"]).dt.date == d) if "Fecha" in festivos_df.columns and not festivos_df.empty else False
    except Exception:
        return False

def buscar_tipico(grupo: str, dia: str):
    r = j_horaria[
        (j_horaria["Grupo"].astype(str) == str(grupo))
        & (j_horaria["Dia"].astype(str).str.lower() == str(dia).lower())
    ]
    if r.empty:
        return None, None, LUNCH_I, LUNCH_F

    hi = _to_time(r.iloc[0].get("HoraInicioLabor"))
    hf = _to_time(r.iloc[0].get("HoraFinLabor"))
    ai = _to_time(r.iloc[0].get("HoraInicioAlmuerzo")) or LUNCH_I
    af = _to_time(r.iloc[0].get("HoraFinAlmuerzo")) or LUNCH_F
    return hi, hf, ai, af

def tope_diario(grupo: str, dia: str):
    r = j_diaria[
        (j_diaria["Grupo"].astype(str) == str(grupo))
        & (j_diaria["Dia"].astype(str).str.lower() == str(dia).lower())
    ]
    if r.empty:
        return None
    v = r.iloc[0]["Horas"]
    try:
        return float(str(v).replace(",", "."))
    except Exception:
        return None

def classify_interval(
    start_dt: datetime, end_dt: datetime,
    tipico_inicio: Optional[time], tipico_fin: Optional[time],
    tope_horas_dia_inicio: Optional[float],
    is_festivo_fn: Callable[[date], bool],
    allows_ordinario_fn: Callable[[date], bool],
    descontar_almuerzo: bool = False,
    alm_i: Optional[time] = None, alm_f: Optional[time] = None,
    tope_restante_map: Optional[Dict[Tuple[Tuple, date], int]] = None,
    festivo_base_map: Optional[Dict[Tuple[Tuple, date], int]] = None,
    lunch_consumed_map: Optional[Dict[Tuple[Tuple, date], bool]] = None,
    tag: Optional[Tuple] = None,
    base_festivo_min_fn: Optional[Callable[[date], int]] = None,
    get_tope_horas_fn: Optional[Callable[[date], Optional[float]]] = None,
) -> MinutesDict:
    mins: MinutesDict = dict(OD=0, ON=0, ED=0, EN=0, FD=0, FN=0, EFD=0, EFN=0)

    if tope_restante_map is None:
        tope_restante_map = {}
    if festivo_base_map is None:
        festivo_base_map = {}
    if lunch_consumed_map is None:
        lunch_consumed_map = {}

    def _ensure_tope_for(fecha: date):
        if get_tope_horas_fn is None:
            return
        k = (tag, fecha)
        if k not in tope_restante_map:
            h = get_tope_horas_fn(fecha)
            if h is not None:
                tope_restante_map[k] = int(h * 60)

    def _get_rem(fecha: date) -> Optional[int]:
        return tope_restante_map.get((tag, fecha), None)

    def _consume_tope(fecha: date, m: int) -> Tuple[int, int]:
        rem = _get_rem(fecha)
        if rem is None:
            return 0, m
        use = min(rem, m)
        tope_restante_map[(tag, fecha)] = rem - use
        return use, m - use

    def _festivo_base_disp(fecha: date) -> int:
        base = int(base_festivo_min_fn(fecha)) if base_festivo_min_fn else 0
        if base <= 0:
            base = int(DEFAULT_BASE_FESTIVA_HORAS * 60)
        usado = festivo_base_map.get((tag, fecha), 0)
        return max(0, base - usado)

    def _consume_festivo_base(fecha: date, m: int) -> Tuple[int, int]:
        disp = _festivo_base_disp(fecha)
        use = min(disp, m)
        if use > 0:
            festivo_base_map[(tag, fecha)] = festivo_base_map.get((tag, fecha), 0) + use
        return use, m - use

    def _can_deduct_lunch(fecha: date) -> bool:
        return not lunch_consumed_map.get((tag, fecha), False)

    def _mark_lunch(fecha: date):
        lunch_consumed_map[(tag, fecha)] = True

    if tope_horas_dia_inicio is not None:
        tope_restante_map.setdefault((tag, start_dt.date()), int(tope_horas_dia_inicio * 60))

    cur = start_dt
    while cur < end_dt:
        d = cur.date()
        d0 = datetime.combine(d, time(0, 0))
        d6 = datetime.combine(d, DAY_06)
        d18 = datetime.combine(d, DAY_18)
        d24 = datetime.combine(d, time(23, 59, 59, 999000)) + timedelta(microseconds=1000)

        windows = [
            ("NOC_00_06", d0, d6),
            ("DIA_06_18", d6, d18),
            ("NOC_18_24", d18, d24),
        ]

        for wname, w0, w1 in windows:
            s, e = clamp_interval(cur, end_dt, w0, w1)
            if not s:
                continue
            dur = minutes_between(s, e)

            alm = 0
            if descontar_almuerzo and alm_i and alm_f and wname == "DIA_06_18" and _can_deduct_lunch(d):
                ai = datetime.combine(d, alm_i)
                af = datetime.combine(d, alm_f)
                a0, a1 = clamp_interval(s, e, ai, af)
                if a0:
                    alm = minutes_between(a0, a1)
                    if alm > 0:
                        _mark_lunch(d)

            eff = max(0, dur - alm)
            if eff == 0:
                continue

            if wname == "DIA_06_18":
                es_festivo = is_festivo_fn(d)
                permite_ordin = allows_ordinario_fn(d)
                if es_festivo:
                    take, extra = _consume_festivo_base(d, eff)
                    mins["FD"] += take
                    if extra > 0:
                        mins["EFD"] += extra
                else:
                    if not permite_ordin:
                        mins["ED"] += eff
                    else:
                        _ensure_tope_for(d)
                        used, over = _consume_tope(d, eff)
                        mins["OD"] += used
                        if over > 0:
                            mins["ED"] += over

            elif wname == "NOC_00_06":
                es_festivo = is_festivo_fn(d)
                permite_ordin = allows_ordinario_fn(d)
                if es_festivo:
                    take, extra = _consume_festivo_base(d, eff)
                    mins["FN"] += take
                    if extra > 0:
                        mins["EFN"] += extra
                else:
                    if not permite_ordin:
                        mins["EN"] += eff
                    else:
                        _ensure_tope_for(d)
                        used, over = _consume_tope(d, eff)
                        mins["ON"] += used
                        if over > 0:
                            mins["EN"] += over

            else:  # NOC_18_24
                es_festivo = is_festivo_fn(d)
                permite_ordin = allows_ordinario_fn(d)
                if es_festivo:
                    take, extra = _consume_festivo_base(d, eff)
                    mins["FN"] += take
                    if extra > 0:
                        mins["EFN"] += extra
                else:
                    if not permite_ordin:
                        mins["EN"] += eff
                    else:
                        _ensure_tope_for(d)
                        used, over = _consume_tope(d, eff)
                        mins["ON"] += used
                        if over > 0:
                            mins["EN"] += over

        cur = datetime.combine(d + timedelta(days=1), time(0, 0))

    return mins

st.subheader("3) Resultado de clasificación")

# Rebuild bags fresh each calculation
tope_restante_map = {}
festivo_base_map = {}
lunch_consumed_map = {}

df_input = st.session_state.nov_df.copy()
if not df_input.empty:
    df_input["__fi__"] = pd.to_datetime(df_input["Fecha y hora de inicio labores"], errors="coerce")
    df_input.sort_values(by=["Nombre del Funcionario", "__fi__"], inplace=True)

rows = []
for idx, row in df_input.iterrows():
    fi = row.get("Fecha y hora de inicio labores")
    ff = row.get("Fecha y hora final labores")
    func = row.get("Nombre del Funcionario")
    esp = row.get("Especialidad")
    grupo = row.get("Grupo")
    equipo = row.get("Equipo")
    actividad = row.get("Actividad / Frente de trabajo")
    sub_dias_row = row.get("Subsidio de Transporte (días)") or 0

    if pd.isna(fi) or pd.isna(ff) or pd.isna(grupo):
        continue

    fi = pd.to_datetime(fi)
    ff = pd.to_datetime(ff)
    if ff <= fi:
        st.error(f"Registro de {func}: 'Fecha fin' debe ser mayor que 'Fecha inicio'.")
        continue

    d = fi.date()
    dia_nom = DIAS_ES[d.weekday()]
    tip_ini, tip_fin, alm_i, alm_f = buscar_tipico(grupo, dia_nom)
    tope_ini_h = tope_diario(grupo, dia_nom)

    def _base_festivo_min(fecha_d: date) -> int:
        dn = DIAS_ES[fecha_d.weekday()]
        th = tope_diario(grupo, dn)
        if th is None:
            th = DEFAULT_BASE_FESTIVA_HORAS
        return int(th * 60)

    def _get_tope_horas(fecha_d: date) -> Optional[float]:
        dn = DIAS_ES[fecha_d.weekday()]
        return tope_diario(grupo, dn)

    def _allows_ordinario(fecha_d: date) -> bool:
        wd = fecha_d.weekday()
        if wd == 5:
            return bool(st.session_state.get("permitir_sabado_ordinario", False))
        if wd == 6:
            return bool(st.session_state.get("permitir_domingo_ordinario", False))
        return True

    tag = (func,)

    k_ini = (tag, d)
    if k_ini not in tope_restante_map and tope_ini_h is not None:
        tope_restante_map[k_ini] = int(tope_ini_h * 60)

    if fi.date() != ff.date():
        d_sig = (fi + timedelta(days=1)).date()
        tope_sig_h = _get_tope_horas(d_sig)
        if tope_sig_h is not None:
            k_sig = (tag, d_sig)
            if k_sig not in tope_restante_map:
                tope_restante_map[k_sig] = int(tope_sig_h * 60)

    mins = classify_interval(
        fi.to_pydatetime(), ff.to_pydatetime(),
        tip_ini, tip_fin,
        None,
        es_festivo_fn_local,
        _allows_ordinario,
        bool(row.get("Descontar Almuerzo", True)), alm_i, alm_f,
        tope_restante_map=tope_restante_map,
        festivo_base_map=festivo_base_map,
        lunch_consumed_map=lunch_consumed_map,
        tag=tag,
        base_festivo_min_fn=_base_festivo_min,
        get_tope_horas_fn=_get_tope_horas
    )

    total_min = sum(mins.values())
    total_h = round(total_min / 60, 2)
    horas_adic = round(
        (mins.get("ED", 0) + mins.get("EN", 0) + mins.get("FD", 0) +
         mins.get("FN", 0) + mins.get("EFD", 0) + mins.get("EFN", 0)) / 60, 2
    )

    tipo_principal = max(mins, key=mins.get) if total_min > 0 else None
    if len([v for v in mins.values() if v == max(mins.values()) and v > 0]) > 1:
        tipo_principal = "Mixta"

    rows.append({
        "Nombre del Funcionario": func,
        "Especialidad": esp,
        "Grupo": grupo,
        "Equipo": equipo,
        "Fecha": d,
        "Día": dia_nom,
        "Mes": fi.strftime("%B").capitalize(),
        "Actividad / Frente de trabajo": actividad,
        "OD (h)": math.ceil(mins["OD"] / 60),
        "ON (h)": math.ceil(mins["ON"] / 60),
        "ED (h)": math.ceil(mins["ED"] / 60),
        "EN (h)": math.ceil(mins["EN"] / 60),
        "FD (h)": math.ceil(mins["FD"] / 60),
        "FN (h)": math.ceil(mins["FN"] / 60),
        "EFD (h)": math.ceil(mins["EFD"] / 60),
        "EFN (h)": math.ceil(mins["EFN"] / 60),
        "Total Horas (todas)": math.ceil(total_h),
        "Total Horas Adicionales": math.ceil(horas_adic),
        "Subsidio de Transporte (días)": sub_dias_row,
    })

df_res = pd.DataFrame(rows)

DEC_COLS = [
    "OD (h)", "ON (h)", "ED (h)", "EN (h)",
    "FD (h)", "FN (h)", "EFD (h)", "EFN (h)",
    "Total Horas (todas)", "Total Horas Adicionales"
]
for c in DEC_COLS:
    if c in df_res.columns:
        df_res[c] = pd.to_numeric(df_res[c], errors="coerce").round(2)

st.subheader("3.1) Tabla de resultados")
if not df_res.empty:
    df_res["Mes"] = pd.to_datetime(df_res["Fecha"]).apply(mes_en_espanol)

st.dataframe(df_res, use_container_width=True)

st.subheader("4) Resúmenes")
if not df_res.empty:
    # Cálculos de resumen
    total_OD = int(df_res["OD (h)"].sum())
    total_ON = int(df_res["ON (h)"].sum())
    total_adic = int( df_res[["ED (h)","EN (h)","FD (h)","FN (h)","EFD (h)","EFN (h)"]].sum(axis=1).sum() )
    total_dias_subsidio = int( pd.to_numeric(df_res["Subsidio de Transporte (días)"], errors="coerce").fillna(0).sum() )
    horas_totales = int(total_OD + total_ON + total_adic)

    colA, colB, colC, colD, colE = st.columns(5)
    with colA:
        st.metric("Horas Totales (Ordinarias + Adicionales)", horas_totales)
    with colB:
        st.metric("OD Totales", total_OD)
    with colC:
        st.metric("ON Totales", total_ON)
    with colD:
        st.metric("Horas Adicionales (ED+EN+FD+FN+EFD+EFN)", total_adic)
    with colE:
        st.metric("Días de Subsidio", total_dias_subsidio)

    cols_sum = ["OD (h)", "ON (h)", "ED (h)", "EN (h)", "FD (h)", "FN (h)", "EFD (h)", "EFN (h)"]
    
    # ---- AJUSTE AQUÍ ----
    # 1. Crea la tabla pivote y resetea el índice
    piv = df_res.pivot_table(index=["Nombre del Funcionario", "Grupo"], values=cols_sum, aggfunc="sum").reset_index()
    
    # 2. Aplica .astype(int) SOLO a las columnas de la suma (cols_sum)
    piv[cols_sum] = piv[cols_sum].astype(int)
    # ---------------------

    st.dataframe(piv, use_container_width=True)



def export_excel(df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm") as writer:
        # Hoja principal
        df.to_excel(writer, index=False, sheet_name="Novedades")
        
        if not df.empty:
            # Columnas a sumar
            cols_sum = ["OD (h)", "ON (h)", "ED (h)", "EN (h)", "FD (h)", "FN (h)", 
                       "EFD (h)", "EFN (h)", "Total Horas (todas)", "Total Horas Adicionales"]
            cols_sum = [c for c in cols_sum if c in df.columns]  # Solo las que existan
            
            # Pivot resumen
            piv = df.pivot_table(
                index=["Nombre del Funcionario", "Grupo"], 
                values=cols_sum, 
                aggfunc="sum"
            ).reset_index()
            piv.to_excel(writer, index=False, sheet_name="Resumen")
            
            # Autoajuste de columnas
            for sheet in ["Novedades", "Resumen"]:
                worksheet = writer.sheets[sheet]
                for i, col in enumerate(df.columns if sheet == "Novedades" else piv.columns):
                    max_len = max(df[col].astype(str).map(len).max() if sheet == "Novedades" else piv[col].astype(str).map(len).max(), len(col))
                    worksheet.set_column(i, i, min(max_len + 2, 50))
    
    return out.getvalue()


# -----------------------------
# 2. Exportar PDF (robusto)
# -----------------------------
def _build_html_report(df: pd.DataFrame) -> str:
    # --- Estilos (agregamos clase para el logo) ---
    estilo = """
    <style>
      @page { size: A4 landscape; margin: 1.2cm; }
      body { font-family: 'DejaVu Sans', Arial, sans-serif; font-size: 10px; color: #222; position: relative; }
      h1 { color: #2b8a3e; font-size: 18px; margin:0 0 6px 0; border-bottom:2px solid #2b8a3e; padding-bottom:5px; }
      h2 { color: #444; font-size: 13px; margin:22px 0 8px 0; }
      .kpi-container { display:flex; flex-wrap:wrap; gap:12px; margin:15px 0; }
      .kpi { background:#eef7f0; padding:8px 14px; border-radius:8px; font-weight:600; font-size:11px; }
      table { width:100%; border-collapse:collapse; margin-top:10px; font-size:9px; page-break-inside:avoid; }
      th { background:#f0f7f2; padding:6px 4px; text-align:center; font-weight:600; border:1px solid #ccc; }
      td { padding:4px 3px; border:1px solid #ddd; vertical-align:top; }
      .generated { font-size:9px; color:#666; margin-top:10px; }
      @media print { body { -webkit-print-color-adjust: exact; } }

      /* MARCA DE AGUA: LOGO EPM desde URL */
      .watermark-logo {
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        width: 500px;
        max-width: 70%;
        opacity: 0.07;
        z-index: -1;
        pointer-events: none;
      }
    </style>
    """

    # --- KPIs ---
    total_OD = int(df["OD (h)"].sum())
    total_ON = int(df["ON (h)"].sum())
    total_adic = int(df[["ED (h)","EN (h)","FD (h)","FN (h)","EFD (h)","EFN (h)"]].sum().sum())
    horas_totales = total_OD + total_ON + total_adic
    dias_sub = int(pd.to_numeric(df["Subsidio de Transporte (días)"], errors="coerce").fillna(0).sum())

    kpis = f"""
    <div class="kpi-container">
      <div class="kpi">Horas Totales: <strong>{horas_totales}</strong></div>
      <div class="kpi">OD: <strong>{total_OD}</strong></div>
      <div class="kpi">ON: <strong>{total_ON}</strong></div>
      <div class="kpi">H. Adicionales: <strong>{total_adic}</strong></div>
      <div class="kpi">Subsidio (días): <strong>{dias_sub}</strong></div>
    </div>
    """

    # --- Tabla ---
    tabla = df.to_html(index=False, border=0, justify="center")

    # --- Fecha en español ---
    meses_es = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    hoy = datetime.now()
    fecha_gen_es = f"{hoy.day} de {meses_es[hoy.month - 1]} de {hoy.year}"

    # --- HTML con logo desde URL ---
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      {estilo}
    </head>
    <body>
      <!-- MARCA DE AGUA: LOGO EPM desde URL -->
      <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/6/62/Logo_EPM.svg/990px-Logo_EPM.svg.png" class="watermark-logo" alt="EPM">

      <div style="position:relative; z-index:1;">
        <h1>Reporte Novedades de Tiempo</h1>
        <p><em>Generado el {fecha_gen_es}</em></p>
        <h2>Resumen General</h2>
        {kpis}
        <h2>Detalle de Registros</h2>
        {tabla}
      </div>
    </body>
    </html>
    """


# ==============================
# EXPORTAR PDF con WEASYPRINT (PROFESIONAL + LOGO DESDE URL)
# ==============================
from io import BytesIO
import pandas as pd
import streamlit as st

# URL del logo EPM (oficial, SVG de alta calidad)
LOGO_EPM_URL = "https://upload.wikimedia.org/wikipedia/commons/thumb/6/62/Logo_EPM.svg/990px-Logo_EPM.svg.png"

def _build_html_report(df: pd.DataFrame) -> str:
    # --- ESTILOS CSS PROFESIONALES + MARCA DE AGUA CON URL ---
    estilo = f"""
    <style>
      @page {{ size: A4 landscape; margin: 1cm; }}
      body {{ 
        font-family: 'DejaVu Sans', Arial, sans-serif; 
        font-size: 10px; 
        color: #222; 
        margin: 0; 
        position: relative;
      }}
      h1 {{ 
        color: #2b8a3e; 
        font-size: 18px; 
        margin-bottom: 5px; 
        border-bottom: 2px solid #2b8a3e; 
        padding-bottom: 5px; 
      }}
      h2 {{ 
        color: #444; 
        font-size: 13px; 
        margin: 20px 0 8px 0; 
      }}
      .header {{ margin-bottom: 15px; }}
      .kpi-container {{ 
        display: flex; 
        flex-wrap: wrap; 
        gap: 12px; 
        margin: 15px 0; 
      }}
      .kpi {{ 
        background: #eef7f0; 
        padding: 8px 14px; 
        border-radius: 8px; 
        font-weight: 600; 
        font-size: 11px; 
      }}
      table {{ 
        width: 100%; 
        border-collapse: collapse; 
        margin-top: 10px; 
        font-size: 9px; 
        page-break-inside: avoid;
      }}
      th {{ 
        background: #f0f7f2; 
        padding: 6px 4px; 
        text-align: center; 
        font-weight: 600; 
        border: 1px solid #ccc; 
      }}
      td {{ 
        padding: 4px 3px; 
        border: 1px solid #ddd; 
        vertical-align: top; 
      }}
      .generated {{ 
        font-size: 9px; 
        color: #666; 
        margin-top: 10px; 
      }}

      /* MARCA DE AGUA: LOGO EPM DESDE URL */
      .watermark-logo {{
        position: fixed;
        top: 50%;
        left: 60%;
        transform: translate(-50%, -50%);
        width: 500px;
        max-width: 65%;
        opacity: 0.07;
        z-index: -1;
        pointer-events: none;
      }}

      @media print {{ 
        body {{ -webkit-print-color-adjust: exact; }} 
      }}
    </style>
    """

    # --- KPIs ---
    total_OD = int(df["OD (h)"].sum())
    total_ON = int(df["ON (h)"].sum())
    total_adic = int(df[["ED (h)","EN (h)","FD (h)","FN (h)","EFD (h)","EFN (h)"]].sum().sum())
    horas_totales = total_OD + total_ON + total_adic
    dias_sub = int(pd.to_numeric(df["Subsidio de Transporte (días)"], errors="coerce").fillna(0).sum())

    kpis = f"""
    <div class="kpi-container">
      <div class="kpi">Horas Totales: <strong>{horas_totales}</strong></div>
      <div class="kpi">OD: <strong>{total_OD}</strong></div>
      <div class="kpi">ON: <strong>{total_ON}</strong></div>
      <div class="kpi">H. Adicionales: <strong>{total_adic}</strong></div>
      <div class="kpi">Subsidio (días): <strong>{dias_sub}</strong></div>
    </div>
    """

    # --- Tabla completa ---
    tabla = df.to_html(index=False, border=0, table_id="data", justify="center")

    # --- Fecha generada ---
    fecha_gen = pd.Timestamp('now').strftime('%d/%m/%Y %H:%M')

    # --- HTML FINAL con logo desde URL ---
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      {estilo}
    </head>
    <body>
      <!-- MARCA DE AGUA: LOGO EPM DESDE URL -->
      <img src="{LOGO_EPM_URL}" class="watermark-logo" alt="EPM">

      <div class="header" style="position: relative; z-index: 1;">
        <h1>Reporte Novedades de Tiempo</h1>
        <div class="generated">Generado: {fecha_gen}</div>
      </div>

      <h2>Resumen General</h2>
      {kpis}

      <h2>Detalle de Registros</h2>
      {tabla}
    </body>
    </html>
    """


def export_pdf_bytes(df: pd.DataFrame) -> tuple[bytes, str]:
    """Devuelve (bytes_pdf, mime_type)"""
    html_content = _build_html_report(df)

    try:
        from weasyprint import HTML
        pdf_buffer = BytesIO()
        HTML(string=html_content, base_url=".").write_pdf(pdf_buffer)
        pdf_buffer.seek(0)
        return pdf_buffer.read(), "application/pdf"
    except ImportError:
        st.warning("WeasyPrint no instalado. Se entrega HTML imprimible.")
        return html_content.encode("utf-8"), "text/html"
    except Exception as e:
        st.error(f"Error generando PDF: {e}")
        return html_content.encode("utf-8"), "text/html"


# -----------------------------
# UI: Botones de Exportación
# -----------------------------
st.subheader("5) Exportar Reporte")

col1, col2 = st.columns(2)

with col1:
    if not df_res.empty:
        excel_bytes = export_excel(df_res)
        st.download_button(
            label="Descargar Excel (.xlsx)",
            data=excel_bytes,
            file_name="Reporte_Novedades_Acumulacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.caption("No hay datos para exportar.")

with col2:
    if not df_res.empty:
        pdf_bytes, mime_type = export_pdf_bytes(df_res)
        ext = "pdf" if mime_type == "application/pdf" else "html"
        label = f"Descargar PDF (.{ext})"
        if ext == "html":
            label += " (imprime como PDF)"
        st.download_button(
            label=label,
            data=pdf_bytes,
            file_name=f"Reporte_Novedades.{ext}",
            mime=mime_type,
            use_container_width=True
        )
    else:
        st.caption("No hay datos para exportar.")
