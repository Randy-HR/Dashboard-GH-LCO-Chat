import pandas as pd
import streamlit as st
import plotly.express as px

# ---------------------------
# Configuración de la página
# ---------------------------
st.set_page_config(page_title="Human Resources Dashboard - Liebherr Colombia SAS", layout="wide")

# ---------------------------
# Lista completa de meses
# ---------------------------
all_months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# -----------------------------------------
# URL del Excel desde Secrets (Streamlit)
# -----------------------------------------
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

# -----------------------------------------
# Sidebar - Botón para refrescar datos
# -----------------------------------------
st.sidebar.title("Filtros")
if st.sidebar.button("🔄 Actualizar datos"):
    st.cache_data.clear()

# -----------------------------------------
# Validación de URL
# -----------------------------------------
if not EXCEL_URL:
    st.error("No se encontró EXCEL_URL en Secrets. Agrega la variable en Settings > Secrets con el vínculo de OneDrive/SharePoint.")
    st.stop()

# -----------------------------------------
# Cargar datos remotos
# -----------------------------------------
try:
    df_total = load_excel_from_url(EXCEL_URL)
except Exception as ex:
    st.error(f"⚠️ No se pudo leer el Excel desde la URL. Verifica permisos del vínculo y que el enlace descargue el archivo.\n\nDetalle: {ex}")
    st.stop()

# -----------------------------------------
# Transformación a formato largo
# -----------------------------------------
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

# -----------------------------------------
# Sidebar - Filtros
# -----------------------------------------
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

# -----------------------------------------
# Selección de tipos de gráfico por indicador
# -----------------------------------------
chart_options = [
    "Bar", "Column", "Line", "Area", "Pie", "Scatter", "Bubble",
    "Box", "Histogram", "Heatmap", "Violin"
]
chart_type_by_indicator = {}
st.sidebar.markdown("### Tipo de gráfico por indicador")
for indicator in selected_indicators:
    chart_type_by_indicator[indicator] = st.sidebar.multiselect(f"{indicator}", chart_options, default=["Bar"]) 

# -----------------------------------------
# Paleta de colores
# -----------------------------------------
palette_option = st.sidebar.selectbox("Selecciona paleta de colores", ["Gris-Amarillo-Negro-Blanco", "Azul-Rojo-Verde", "Pastel"]) 
if palette_option == "Gris-Amarillo-Negro-Blanco":
    color_sequence = ["gray", "yellow", "black", "white"]
elif palette_option == "Azul-Rojo-Verde":
    color_sequence = ["blue", "red", "green"]
else:
    color_sequence = px.colors.qualitative.Pastel

# -----------------------------------------
# Título y tablero
# -----------------------------------------
st.title("Human Resources Dashboard - Liebherr Colombia SAS")
st.markdown("## Visualización de Indicadores")
cols = st.columns(4)
chart_count = 0
max_charts = 12  # mantener 12 como solicitaste

for indicator in selected_indicators:
    indicator_df = filtered_df[filtered_df["Indicator Name"] == indicator]
    if indicator_df.empty:
        continue
    for chart_type in chart_type_by_indicator.get(indicator, []):
        fig = None
        if chart_type == "Bar":
            fig = px.bar(indicator_df, x="Month", y="Value", color="Year", barmode="group", title=f"{indicator} - Bar Chart", color_discrete_sequence=color_sequence)
        elif chart_type == "Column":
            fig = px.bar(indicator_df, x="Year", y="Value", color="Month", barmode="group", title=f"{indicator} - Column Chart", color_discrete_sequence=color_sequence)
        elif chart_type == "Line":
            fig = px.line(indicator_df, x="Month", y="Value", color="Year", title=f"{indicator} - Line Chart", color_discrete_sequence=color_sequence)
        elif chart_type == "Area":
            fig = px.area(indicator_df, x="Month", y="Value", color="Year", title=f"{indicator} - Area Chart", color_discrete_sequence=color_sequence)
        elif chart_type == "Pie":
            pie_df = indicator_df.groupby("Month")["Value"].sum().reset_index()
            fig = px.pie(pie_df, names="Month", values="Value", hole=0.4, title=f"{indicator} - Donut Chart", color_discrete_sequence=color_sequence)
        elif chart_type == "Scatter":
            fig = px.scatter(indicator_df, x="Month", y="Value", color="Year", title=f"{indicator} - Scatter Plot", color_discrete_sequence=color_sequence)
        elif chart_type == "Bubble":
            fig = px.scatter(indicator_df, x="Month", y="Value", size="Value", color="Year", title=f"{indicator} - Bubble Chart", color_discrete_sequence=color_sequence)
        elif chart_type == "Box":
            fig = px.box(indicator_df, x="Month", y="Value", color="Year", title=f"{indicator} - Box Plot", color_discrete_sequence=color_sequence)
        elif chart_type == "Histogram":
            fig = px.histogram(indicator_df, x="Value", color="Year", title=f"{indicator} - Histogram", color_discrete_sequence=color_sequence)
        elif chart_type == "Heatmap":
            heatmap_df = indicator_df.pivot_table(index="Month", columns="Year", values="Value", aggfunc="sum")
            fig = px.imshow(heatmap_df, text_auto=True, title=f"{indicator} - Heatmap")
        elif chart_type == "Violin":
            fig = px.violin(indicator_df, x="Month", y="Value", color="Year", box=True, title=f"{indicator} - Violin Plot", color_discrete_sequence=color_sequence)

        if fig:
            cols[chart_count % 4].plotly_chart(fig, use_container_width=True)
            chart_count += 1
            if chart_count >= max_charts:
                break
    if chart_count >= max_charts:
        break

# ==================================================
# 💬 Asistente de IA (LangChain + Azure OpenAI)
# ==================================================
from langchain_experimental.agents import create_pandas_dataframe_agent
from langchain_openai import AzureChatOpenAI

SYSTEM_GUIDE = """
Eres un analista senior de Gestión Humana. Responde SIEMPRE apoyándote en cálculos sobre el DataFrame disponible (filtered_df).
- Si te piden “mayor/menor”, calcula máximos/mínimos y referencia mes y año.
- Para “tendencias” o “proyecciones” usa una regresión lineal simple (numpy.polyfit) sobre la serie solicitada y aclara supuestos.
- Devuelve tablas concisas cuando aporten claridad; redondea a 2 decimales.
- No leas ni escribas archivos, no uses red, no ejecutes código fuera de pandas/numpy.
- Si el indicador/columna no existe, explica cómo encontrarlo en el dashboard.
"""

# Historial de chat (una única caja de chat en toda la página)
if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = [
        {"role": "assistant", "content": "Hola, soy tu asistente de análisis. Pregúntame por máximos, mínimos, comparaciones o proyecciones sobre los indicadores filtrados."}
    ]

# Muestra historial
for m in st.session_state.chat_messages:
    with st.chat_message(m["role"]):
        st.markdown(m["content"])

# Constructor del LLM de Azure OpenAI
def _build_llm():
    try:
        return AzureChatOpenAI(
            azure_endpoint=st.secrets["AZURE_OPENAI_ENDPOINT"],
            api_version=st.secrets.get("AZURE_OPENAI_API_VERSION", "2024-02-01"),
            api_key=st.secrets["AZURE_OPENAI_API_KEY"],
            deployment_name=st.secrets["AZURE_OPENAI_DEPLOYMENT"],
            temperature=0.0,
        )
    except Exception:
        st.error("Faltan Secrets de Azure OpenAI (AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_API_KEY, AZURE_OPENAI_DEPLOYMENT).")
        st.stop()

@st.cache_resource(show_spinner=False)
def _build_agent_for_df(sample_df: pd.DataFrame):
    llm = _build_llm()
    agent = create_pandas_dataframe_agent(
        llm,
        sample_df,                # trabajamos sobre el DF filtrado
        verbose=False,
        max_iterations=5,
        include_df_in_prompt=False
    )
    return agent

# Caja de chat (solo una por página)
user_q = st.chat_input("Escribe aquí tu pregunta (p. ej.: '¿Mes con mayor Fluctuation Rate I en 2025?' o 'Proyecta Cost per hire para Q1 2026').")
if user_q:
    st.session_state.chat_messages.append({"role": "user", "content": user_q})
    with st.chat_message("user"):
        st.markdown(user_q)

    with st.chat_message("assistant"):
        with st.status("Analizando datos del dashboard...", expanded=False):
            try:
                agent = _build_agent_for_df(filtered_df)
                full_q = f"{SYSTEM_GUIDE}\n\nPregunta: {user_q}\n\nRecuerda: el DataFrame disponible se llama 'filtered_df' y contiene las columnas: {list(filtered_df.columns)}."
                answer = agent.run(full_q)
            except Exception as ex:
                answer = ("⚠️ No pude completar el análisis.\n"
                          f"Detalle técnico: {ex}\n"
                          "Revisa que el indicador exista y que los filtros del dashboard contengan datos.")
        st.markdown(answer)
    st.session_state.chat_messages.append({"role": "assistant", "content": answer})
