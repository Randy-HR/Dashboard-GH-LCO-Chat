import pandas as pd
import streamlit as st
import plotly.express as px

# -------------------------------
# Configuración de la página
# -------------------------------
st.set_page_config(page_title="Human Resources Dashboard - Liebherr Colombia SAS", layout="wide")

# -------------------------------
# Lista completa de meses
# -------------------------------
all_months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# -------------------------------
# URL del Excel desde Secrets (Streamlit)
# -------------------------------
EXCEL_URL = st.secrets.get("EXCEL_URL")

@st.cache_data(ttl=300)
def load_excel_from_url(url: str) -> pd.DataFrame:
    """Lee el libro de Excel desde una URL (OneDrive/SharePoint) y concatena todas las hojas.
    Requiere que el enlace sirva descarga directa del .xlsx (ej.: &download=1).
    """
    xls = pd.ExcelFile(url, engine="openpyxl")
    df_list = []
    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)
        df["Year"] = sheet_name
        df_list.append(df)
    df_total = pd.concat(df_list, ignore_index=True)
    return df_total

# -------------------------------
# Sidebar - Botón para refrescar datos
# -------------------------------
st.sidebar.title("Filtros")
if st.sidebar.button("🔄 Actualizar datos"):
    st.cache_data.clear()

# -------------------------------
# Validación de URL
# -------------------------------
if not EXCEL_URL:
    st.error("No se encontró EXCEL_URL en Secrets. Agrega la variable en Settings > Secrets con el vínculo de OneDrive/SharePoint.")
    st.stop()

# -------------------------------
# Cargar datos remotos
# -------------------------------
try:
    df_total = load_excel_from_url(EXCEL_URL)
except Exception as ex:
    st.error(f"⚠️ No se pudo leer el Excel desde la URL. Verifica permisos del vínculo y que el enlace descargue el archivo.\n\nDetalle: {ex}")
    st.stop()

# -------------------------------
# Transformación a formato largo
# -------------------------------
id_cols = [
    "Subprocess", "Indicator Type", "Indicator Name",
    "Measurement Frequency", "Formula / How it is calculated", "Year"
]
missing_cols = [c for c in id_cols if c not in df_total.columns]
if missing_cols:
    st.error(f"Faltan columnas requeridas en el Excel: {missing_cols}. Verifica el esquema del archivo.")
    st.stop()

present_months = [m for m in all_months if m in df_total.columns]
if not present_months:
    st.error("No se encontraron columnas de meses (January..December) en el Excel.")
    st.stop()

# Melt con los meses presentes
df_melted = df_total.melt(
    id_vars=id_cols,
    value_vars=present_months,
    var_name="Month",
    value_name="Value"
)

# -------------------------------
# Sidebar - Filtros
# -------------------------------
if st.sidebar.checkbox("Select All Years"):
    selected_years = sorted(df_melted["Year"].unique())
else:
    selected_years = st.sidebar.multiselect("Selecciona Año(s)", sorted(df_melted["Year"].unique()))
filtered_df = df_melted[df_melted["Year"].isin(selected_years)] if selected_years else df_melted

if st.sidebar.checkbox("Select All Months"):
    selected_months = present_months
else:
    selected_months = st.sidebar.multiselect("Selecciona Mes(es)", present_months)
filtered_df = filtered_df[filtered_df["Month"].isin(selected_months)] if selected_months else filtered_df

if st.sidebar.checkbox("Select All Subprocesses"):
    selected_subprocesses = sorted(filtered_df["Subprocess"].dropna().unique())
else:
    selected_subprocesses = st.sidebar.multiselect("Selecciona Subprocess", sorted(filtered_df["Subprocess"].dropna().unique()))
filtered_df = filtered_df[filtered_df["Subprocess"].isin(selected_subprocesses)] if selected_subprocesses else filtered_df

available_indicators = sorted(filtered_df["Indicator Name"].dropna().unique())
if st.sidebar.checkbox("Select All Indicators"):
    selected_indicators = available_indicators
else:
    selected_indicators = st.sidebar.multiselect("Selecciona Indicator(s)", available_indicators)
filtered_df = filtered_df[filtered_df["Indicator Name"].isin(selected_indicators)] if selected_indicators else filtered_df

# -------------------------------
# Selección de tipos de gráfico por indicador
# -------------------------------
chart_options = [
    "Bar", "Column", "Line", "Area", "Pie", "Scatter", "Bubble",
    "Box", "Histogram", "Heatmap", "Violin"
]
chart_type_by_indicator = {}
st.sidebar.markdown("### Tipo de gráfico por indicador")
for indicator in selected_indicators:
    chart_type_by_indicator[indicator] = st.sidebar.multiselect(f"{indicator}", chart_options, default=["Bar"])

# -------------------------------
# Paleta de colores
# -------------------------------
palette_option = st.sidebar.selectbox("Selecciona paleta de colores", ["Gris-Amarillo-Negro-Blanco", "Azul-Rojo-Verde", "Pastel"])
if palette_option == "Gris-Amarillo-Negro-Blanco":
    color_sequence = ["gray", "yellow", "black", "white"]
elif palette_option == "Azul-Rojo-Verde":
    color_sequence = ["blue", "red", "green"]
else:
    color_sequence = px.colors.qualitative.Pastel

# -------------------------------
# Título y tablero
# -------------------------------
st.title("Human Resources Dashboard - Liebherr Colombia SAS")
st.markdown("## Visualización de Indicadores")
cols = st.columns(4)
chart_count = 0
max_charts = 12  # mantener 12

for indicator in selected_indicators:
    indicator_df = filtered_df[filtered_df["Indicator Name"] == indicator]
    if indicator_df.empty:
        continue
    for chart_type in chart_type_by_indicator.get(indicator, []):
        fig = None
        if chart_type == "Bar":
            fig = px.bar(indicator_df, x="Month", y="Value", color="Year", barmode="group",
                         title=f"{indicator} - Bar Chart", color_discrete_sequence=color_sequence)
        elif chart_type == "Column":
            fig = px.bar(indicator_df, x="Year", y="Value", color="Month", barmode="group",
                         title=f"{indicator} - Column Chart", color_discrete_sequence=color_sequence)
        elif chart_type == "Line":
            fig = px.line(indicator_df, x="Month", y="Value", color="Year",
                          title=f"{indicator} - Line Chart", color_discrete_sequence=color_sequence)
        elif chart_type == "Area":
            fig = px.area(indicator_df, x="Month", y="Value", color="Year",
                          title=f"{indicator} - Area Chart", color_discrete_sequence=color_sequence)
        elif chart_type == "Pie":
            pie_df = indicator_df.groupby("Month")["Value"].sum().reset_index()
            fig = px.pie(pie_df, names="Month", values="Value", hole=0.4,
                         title=f"{indicator} - Donut Chart", color_discrete_sequence=color_sequence)
        elif chart_type == "Scatter":
            fig = px.scatter(indicator_df, x="Month", y="Value", color="Year",
                             title=f"{indicator} - Scatter Plot", color_discrete_sequence=color_sequence)
        elif chart_type == "Bubble":
            fig = px.scatter(indicator_df, x="Month", y="Value", size="Value", color="Year",
                             title=f"{indicator} - Bubble Chart", color_discrete_sequence=color_sequence)
        elif chart_type == "Box":
            fig = px.box(indicator_df, x="Month", y="Value", color="Year",
                         title=f"{indicator} - Box Plot", color_discrete_sequence=color_sequence)
        elif chart_type == "Histogram":
            fig = px.histogram(indicator_df, x="Value", color="Year",
                               title=f"{indicator} - Histogram", color_discrete_sequence=color_sequence)
        elif chart_type == "Heatmap":
            heatmap_df = indicator_df.pivot_table(index="Month", columns="Year", values="Value", aggfunc="sum")
            fig = px.imshow(heatmap_df, text_auto=True, title=f"{indicator} - Heatmap")
        elif chart_type == "Violin":
            fig = px.violin(indicator_df, x="Month", y="Value", color="Year", box=True,
                            title=f"{indicator} - Violin Plot", color_discrete_sequence=color_sequence)
        if fig is not None:
            cols[chart_count % 4].plotly_chart(fig, use_container_width=True)
            chart_count += 1
            if chart_count >= max_charts:
                break
    if chart_count >= max_charts:
        break

# ==================================================
# 🗨️ Chat analítico sin IA (reglas)
# ==================================================
import re
import numpy as np

st.markdown("## 🗨️ Chat analítico (sin IA)")
st.caption("Haz preguntas con palabras como: máximo, mínimo, promedio, valor más alto, valor más bajo, o pide un 'resumen ejecutivo'.")

# --- Utilidades de parsing ---
def _normalize(text: str) -> str:
    return text.lower().strip()

def _parse_years_from_text(text: str, all_years: list[str]) -> list[str]:
    """
    Detecta años de 4 dígitos en el texto y los cruza con los años disponibles.
    Devuelve lista de años (como strings). Si no hay, retorna [] para usar filtros activos.
    """
    candidates = re.findall(r"\b(20\d{2})\b", text)
    years = [y for y in candidates if y in set(map(str, all_years))]
    return years

def _find_indicators_in_text(text: str, indicators: list[str]) -> list[str]:
    """Devuelve los indicadores cuyo nombre aparece (parcial/case-insensitive) en la pregunta."""
    t = _normalize(text)
    hits = [name for name in indicators if _normalize(name) in t]
    if hits:
        return hits
    # Búsqueda por tokens
    tokens = [tok for tok in re.split(r"[^a-zA-Z0-9%]+", t) if tok]
    hits2 = []
    for name in indicators:
        nrm = _normalize(name)
        for tok in tokens:
            if tok and tok in nrm:
                hits2.append(name)
                break
    # únicos
    return list(dict.fromkeys(hits2))

def _value_to_float(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")

def _subset_for_question(df_melted: pd.DataFrame, filtered_df: pd.DataFrame, years_from_q: list[str]) -> pd.DataFrame:
    """
    Si el usuario menciona años explícitos, usamos df_melted con esos años (todos los meses).
    Si no menciona años, usamos el filtered_df (respeta los filtros activos).
    """
    base = df_melted
    if years_from_q:
        base = base[base["Year"].astype(str).isin(years_from_q)]
    else:
        base = filtered_df
    return base.copy()

def _compute_answer(stat: str, df: pd.DataFrame, indicator: str) -> tuple[str, pd.DataFrame]:
    """
    Calcula máximo/mínimo/promedio para un indicador.
    Retorna (mensaje, tabla_df) donde tabla_df puede ser None si no aplica.
    """
    sdf = df[df["Indicator Name"] == indicator].copy()
    if sdf.empty:
        return (f"❌ No hay datos para el indicador **{indicator}** con los filtros/años indicados.", None)

    sdf["Value_num"] = _value_to_float(sdf["Value"])
    sdf = sdf.dropna(subset=["Value_num"])
    if sdf.empty:
        return (f"❌ Los valores del indicador **{indicator}** no son numéricos o están vacíos.", None)

    if stat == "max":
        idx = sdf["Value_num"].idxmax()
        row = sdf.loc[idx]
        msg = (
            f"🔺 **Máximo** de **{indicator}**: **{row['Value_num']:.2f}** "
            f"en **{row['Month']} ({row['Year']})**."
        )
        tbl = sdf.groupby("Year")["Value_num"].mean().reset_index().rename(columns={"Value_num": "Average"})
        tbl["Average"] = tbl["Average"].round(2)
        return (msg, tbl)

    if stat == "min":
        idx = sdf["Value_num"].idxmin()
        row = sdf.loc[idx]
        msg = (
            f"🔻 **Mínimo** de **{indicator}**: **{row['Value_num']:.2f}** "
            f"en **{row['Month']} ({row['Year']})**."
        )
        tbl = sdf.groupby("Year")["Value_num"].mean().reset_index().rename(columns={"Value_num": "Average"})
        tbl["Average"] = tbl["Average"].round(2)
        return (msg, tbl)

    if stat == "avg":
        overall = sdf["Value_num"].mean()
        msg = f"📈 **Promedio** de **{indicator}**: **{overall:.2f}** considerando meses/años del alcance solicitado."
        tbl = sdf.groupby("Year")["Value_num"].mean().reset_index().rename(columns={"Value_num": "Average"})
        tbl["Average"] = tbl["Average"].round(2)
        return (msg, tbl)

    return ("❓ No reconocí la operación solicitada (máximo/mínimo/promedio).", None)

def _exec_summary(df: pd.DataFrame) -> tuple[str, pd.DataFrame, pd.DataFrame]:
    """
    Resumen ejecutivo sobre el alcance actual (df).
    Devuelve mensaje y dos tablas: Top picos globales y promedio por indicador.
    """
    if df.empty:
        return ("❌ No hay datos en el alcance actual para generar resumen.", None, None)

    df2 = df.copy()
    df2["Value_num"] = _value_to_float(df2["Value"])
    df2 = df2.dropna(subset=["Value_num"])

    # Top 5 picos globales
    top_peaks = (
        df2.sort_values("Value_num", ascending=False)
           .loc[:, ["Indicator Name", "Year", "Month", "Value_num"]]
           .head(5)
           .rename(columns={"Value_num": "Value"})
    )
    top_peaks["Value"] = top_peaks["Value"].round(2)

    # Promedio por indicador (Top 5)
    by_indicator_avg = (
        df2.groupby("Indicator Name")["Value_num"]
           .mean()
           .reset_index()
           .rename(columns={"Value_num": "Average"})
           .sort_values("Average", ascending=False)
           .head(5)
    )
    by_indicator_avg["Average"] = by_indicator_avg["Average"].round(2)

    years_in_scope = ", ".join(sorted(map(str, df2["Year"].astype(str).unique())))
    msg = (
        "🧭 **Resumen ejecutivo**\n"
        f"- Años en alcance: **{years_in_scope}**\n"
        f"- Indicadores considerados: **{df2['Indicator Name'].nunique()}**\n"
        f"- Registros (mes-año): **{len(df2)}**\n"
        "Se muestran los **5 picos globales** y el **Top 5 de promedios por indicador**."
    )
    return (msg, top_peaks, by_indicator_avg)

# --- Palabras clave e intención ---
def _detect_intent(text: str) -> str:
    t = _normalize(text)
    if "resumen ejecutivo" in t or (("resumen" in t) and ("ejecutivo" in t)):
        return "summary"
    if ("promedio" in t) or ("media" in t) or ("average" in t):
        return "avg"
    if ("máximo" in t) or ("maximo" in t) or ("mayor" in t) or ("valor más alto" in t) or ("valor mas alto" in t) or ("alto" in t):
        return "max"
    if ("mínimo" in t) or ("minimo" in t) or ("menor" in t) or ("valor más bajo" in t) or ("valor mas bajo" in t) or ("bajo" in t):
        return "min"
    return "unknown"

# --- Historial del chat en sesión ---
if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = [
        {"role": "assistant", "content": "Hola, soy tu asistente analítico sin IA. Pregúntame por **máximo**, **mínimo**, **promedio** o pide un **resumen ejecutivo** sobre los indicadores."}
    ]

# Muestra historial
for m in st.session_state.chat_messages:
    with st.chat_message(m["role"]):
        st.markdown(m["content"])

# Entrada de usuario
user_q = st.chat_input("Escribe aquí tu pregunta (ej.: '¿Máximo de Fluctuation Rate I en 2025?' o 'Resumen ejecutivo').")
if user_q:
    st.session_state.chat_messages.append({"role": "user", "content": user_q})
    with st.chat_message("user"):
        st.markdown(user_q)

    with st.chat_message("assistant"):
        with st.status("Analizando datos del dashboard...", expanded=False):
            try:
                intent = _detect_intent(user_q)
                years_all = sorted(df_melted["Year"].astype(str).unique())
                years_from_q = _parse_years_from_text(user_q, years_all)

                # Elegir el alcance (todos los meses y años mencionados; si no hay años, respetar filtros)
                scope_df = _subset_for_question(df_melted, filtered_df, years_from_q)

                if intent == "summary":
                    msg, tbl1, tbl2 = _exec_summary(scope_df)
                    st.markdown(msg)
                    if tbl1 is not None:
                        st.markdown("**Top 5 picos globales**")
                        st.dataframe(tbl1, use_container_width=True)
                    if tbl2 is not None:
                        st.markdown("**Top 5 promedios por indicador**")
                        st.dataframe(tbl2, use_container_width=True)

                elif intent in {"max", "min", "avg"}:
                    # Determinar indicador
                    indicators_all = sorted(scope_df["Indicator Name"].dropna().unique())
                    hits = _find_indicators_in_text(user_q, indicators_all)

                    # Si no hubo coincidencias en texto, intentar usar los seleccionados en la barra lateral
                    if not hits and selected_indicators:
                        if len(selected_indicators) == 1:
                            hits = [selected_indicators[0]]
                        else:
                            st.info("🔎 No detecté el indicador en tu pregunta. Especifica uno (ej.: '... de Cost per hire').")
                            st.stop()

                    if not hits:
                        st.info("🔎 No detecté el indicador en tu pregunta. Ejemplo: 'Máximo de **Fluctuation Rate I** en 2025'.")
                        st.stop()

                    indicator = hits[0]
                    msg, tbl = _compute_answer(intent, scope_df, indicator)
                    st.markdown(msg)
                    if tbl is not None and not tbl.empty:
                        st.markdown("**Promedio por año (tabla de apoyo)**")
                        st.dataframe(tbl, use_container_width=True)

                else:
                    st.info("❓ No reconocí la intención. Usa **máximo**, **mínimo**, **promedio** o **resumen ejecutivo** y menciona el **indicador** si aplica.")
            except Exception as ex:
                st.error(f"⚠️ Ocurrió un error procesando la consulta.\n\nDetalle: {ex}")
