import streamlit as st
import pandas as pd
from datetime import timedelta
from math import ceil
from io import BytesIO

st.set_page_config(page_title="Plan de Fabricación", layout="wide")
st.title("🧱 Plan de Fabricación de Yesería")

meses_es = {
    'January': 'enero', 'February': 'febrero', 'March': 'marzo',
    'April': 'abril', 'May': 'mayo', 'June': 'junio',
    'July': 'julio', 'August': 'agosto', 'September': 'septiembre',
    'October': 'octubre', 'November': 'noviembre', 'December': 'diciembre'
}

@st.cache_data
def cargar_datos():
    try:
        df1 = pd.read_excel("PRODUCCION KARDEX - DILAN.xlsx", sheet_name="BASE1",
                            usecols=["LINEA", "MOLDE", "CANTIDAD FABRICAR", "STOCK TOTAL",
                                     "1 CAMBIO", "2 CAMBIO", "3 CAMBIO"])
        df2 = pd.read_excel("PRODUCCION KARDEX - DILAN.xlsx", sheet_name="BASE2",
                            usecols=["MOLDE", "MOLDE 1 PERSONA (8 horas)"])
    except Exception as e:
        st.error(f"Error al cargar datos: {e}")
        return pd.DataFrame(), pd.DataFrame()

    for col in ["CANTIDAD FABRICAR", "STOCK TOTAL"]:
        if col in df1.columns:
            df1[col] = pd.to_numeric(df1[col], errors="coerce").fillna(0)

    for col in ["1 CAMBIO", "2 CAMBIO", "3 CAMBIO"]:
        if col in df1.columns:
            df1[col] = pd.to_datetime(df1[col], errors="coerce")

    if "MOLDE" in df2.columns:
        df2["MOLDE"] = df2["MOLDE"].astype(str).str.strip().str.upper()

    if "MOLDE" in df1.columns:
        df1["MOLDE"] = df1["MOLDE"].astype(str).str.strip().str.upper()

    if "MOLDE 1 PERSONA (8 horas)" in df2.columns:
        df2["MOLDE 1 PERSONA (8 horas)"] = pd.to_numeric(df2["MOLDE 1 PERSONA (8 horas)"], errors="coerce")

    df1["POR FABRICAR"] = (df1["CANTIDAD FABRICAR"] - df1["STOCK TOTAL"]).clip(lower=0)

    return df1, df2

df, base2 = cargar_datos()
if df.empty:
    st.stop()

def expandir_fechas(df):
    filas_expandidas = []
    for _, row in df.iterrows():
        for col in ["1 CAMBIO", "2 CAMBIO", "3 CAMBIO"]:
            fecha = row[col]
            if pd.notna(fecha):
                fila_nueva = row.copy()
                fila_nueva["CAMBIO PROGRAMADO"] = fecha
                filas_expandidas.append(fila_nueva)
    return pd.DataFrame(filas_expandidas)

df_exp = expandir_fechas(df)
if df_exp.empty:
    st.warning("No hay datos con fechas de cambio válidas.")
    st.stop()

anios_disponibles = sorted(df_exp["CAMBIO PROGRAMADO"].dt.year.unique())
meses_disponibles = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

col1, col2 = st.columns(2)
with col1:
    anio = st.selectbox("📅 Selecciona el AÑO:", ["Todos"] + anios_disponibles)
with col2:
    mes_num = st.selectbox("📅 Selecciona el MES:", ["Todos"] + list(meses_disponibles.keys()),
                           format_func=lambda x: x if x == "Todos" else meses_disponibles[x])

df_filtrado = df_exp.copy()
if anio != "Todos":
    df_filtrado = df_filtrado[df_filtrado["CAMBIO PROGRAMADO"].dt.year == anio]
if mes_num != "Todos":
    df_filtrado = df_filtrado[df_filtrado["CAMBIO PROGRAMADO"].dt.month == mes_num]

molde_seleccionado = st.selectbox("🔎 Ver datos de un MOLDE específico (opcional):", [""] + sorted(df_filtrado["MOLDE"].dropna().unique()))
if molde_seleccionado:
    df_filtrado = df_filtrado[df_filtrado["MOLDE"] == molde_seleccionado]

linea_seleccionada = st.selectbox("🔎 Filtrar por LÍNEA (opcional):", [""] + sorted(df_filtrado["LINEA"].dropna().unique()))
if linea_seleccionada:
    df_filtrado = df_filtrado[df_filtrado["LINEA"] == linea_seleccionada]

df_filtrado = df_filtrado.sort_values("CAMBIO PROGRAMADO")

if df_filtrado.empty:
    st.warning("No hay datos que coincidan con los filtros seleccionados.")
    st.stop()

fechas_unicas = df_filtrado["CAMBIO PROGRAMADO"].dt.date.unique()

plan_trabajo = []

fechas = sorted(df_filtrado["CAMBIO PROGRAMADO"].dt.date.unique())

for fecha in fechas:
    df_fecha = df_filtrado[df_filtrado["CAMBIO PROGRAMADO"].dt.date == fecha]

    # Clasificación por prioridad
    prioridad_1 = df_fecha[df_fecha["MOLDE"].str.startswith("MYIFZ")]
    prioridad_2 = df_fecha[df_fecha["MOLDE"].str.startswith("MYOP")]
    otros = df_fecha[~df_fecha["MOLDE"].str.startswith(("MYIFZ", "MYOP"))]

    orden_final = pd.concat([prioridad_1, prioridad_2, otros], ignore_index=True)

    dias_ocupados = 0  # para ir adelantando según prioridad

    for _, fila in orden_final.iterrows():
        molde = fila["MOLDE"]
        cantidad = fila["POR FABRICAR"]
        fecha_cambio = fila["CAMBIO PROGRAMADO"]
        fecha_base = fecha_cambio - timedelta(days=2 + dias_ocupados)

        if cantidad == 0:
            instruccion = f"❌ No hay que fabricar el molde {molde}."
        else:
            fila_base2 = base2[base2["MOLDE"] == molde]
            if not fila_base2.empty:
                productividad = fila_base2.iloc[0]["MOLDE 1 PERSONA (8 horas)"]
                if productividad > 0:
                    piezas_dia = productividad * 3
                    dias_necesarios = ceil(cantidad / piezas_dia)
                    fecha_inicio = fecha_base - timedelta(days=dias_necesarios)
                    dias_ocupados += dias_necesarios

                    fecha_inicio_str = fecha_inicio.strftime("%#d de %B de %Y")
                    for eng, esp in meses_es.items():
                        fecha_inicio_str = fecha_inicio_str.replace(eng, esp)
                    fecha_inicio_str = fecha_inicio_str.capitalize()
                    instruccion = f"✅ Empezar a fabricar el {fecha_inicio_str}."
                else:
                    instruccion = f"⚠️ Productividad inválida."
            else:
                instruccion = f"🔔 Molde no registrado en BASE2."

        plan_trabajo.append({
            "FECHA CAMBIO": fecha_cambio.date(),
            "LINEA": fila["LINEA"],
            "MOLDE": molde,
            "CANTIDAD FABRICAR": fila["CANTIDAD FABRICAR"],
            "STOCK TOTAL": fila["STOCK TOTAL"],
            "POR FABRICAR": cantidad,
            "INSTRUCCIÓN": instruccion
        })

# Mostrar la tabla combinada
df_plan = pd.DataFrame(plan_trabajo)
df_plan = df_plan.sort_values("FECHA CAMBIO")
st.dataframe(df_plan, use_container_width=True)

def generar_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Plan de Fabricación", index=False)
    return output.getvalue()

if plan_trabajo:
    df_plan = pd.DataFrame(plan_trabajo)
    st.download_button(
        label="📥 Descargar Plan de Trabajo en Excel",
        data=generar_excel(df_plan),
        file_name="Plan_de_Fabricacion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

total = df_filtrado["POR FABRICAR"].sum()
titulo_total = "toda la planificación"
if anio != "Todos" and mes_num != "Todos":
    titulo_total = f"{meses_disponibles[mes_num]} {anio}"
elif anio != "Todos":
    titulo_total = f"año {anio}"
elif mes_num != "Todos":
    titulo_total = f"mes {meses_disponibles[mes_num]}"

st.info(f"🔢 Total por fabricar en {titulo_total}: **{total:.0f} moldes**")