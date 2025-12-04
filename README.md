# Dashboard de Gestión Humana (Streamlit) — Chat Analítico sin IA

Este proyecto implementa un **dashboard interactivo en Streamlit** para visualizar indicadores de Gestión Humana y hacer consultas en lenguaje natural **sin usar un modelo de IA** (chat basado en reglas).  
Los datos se leen desde un **Excel en OneDrive/SharePoint** y el dashboard puede desplegarse en **Streamlit Cloud**.

---

## ✅ Características

- **Visualización dinámica**: hasta **12 gráficos** simultáneos (barras, líneas, área, pastel, dispersión, histograma, mapa de calor, etc.).
- **Filtros interactivos**: por **Año**, **Mes**, **Subprocess** e **Indicator Name**.
- **Conexión con OneDrive/SharePoint**: lectura directa del Excel (todas las hojas se tratan como años).
- **Chat analítico sin IA**:
  - Entiende palabras clave: **máximo**, **mínimo**, **promedio**, **valor más alto**, **valor más bajo**, **resumen ejecutivo**.
  - Si mencionas **años** (ej. `2024`, `2025`), usa todos los meses de esos años desde el **histórico completo**.  
    Si **no** mencionas años, respeta **los filtros activos** del sidebar.
  - Detecta el **indicador** por coincidencia parcial del nombre y devuelve **mes, año y valor**.  
  - **Resumen ejecutivo**: años en alcance, # indicadores, # registros, **Top 5 picos globales** y **Top 5 promedios por indicador**.

---

## 📁 Estructura del repositorio
