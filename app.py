import streamlit as st
import pandas as pd
import plotly.express as px
import folium
from folium.plugins import HeatMap
from streamlit_folium import st_folium
from fpdf import FPDF
import requests
from io import BytesIO

# Configuración de la página
st.set_page_config(layout="wide", page_title="Panel Vendedores Externos", page_icon="📍")

# ID del archivo de Google Drive (extraído de la URL compartida)
FILE_ID = "1Pe0W62WZff6MHkQs_MnY4bu6E4NMAR7U"
GD_URL = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx"

# Función para cargar datos desde Google Drive
@st.cache_data(ttl=3600)  # Cache por 1 hora
def load_data():
    try:
        response = requests.get(GD_URL)
        response.raise_for_status()  # Lanza error si la descarga falla
        excel_file = BytesIO(response.content)
        
        viajes = pd.read_excel(excel_file, sheet_name="VIAJES")
        formularios = pd.read_excel(excel_file, sheet_name="FORMULARIO")
        
        return viajes, formularios
    except Exception as e:
        st.error(f"Error al cargar datos desde Google Drive: {str(e)}")
        st.stop()

# Carga de datos
viajes, formularios = load_data()

# Preprocesamiento fechas
viajes["FECHA"] = pd.to_datetime(viajes["FECHA"])
formularios["Fecha de llenar"] = pd.to_datetime(formularios["Fecha de llenar"])
formularios["Fecha_dia"] = formularios["Fecha de llenar"].dt.date

# Sidebar
st.sidebar.title("📅 Filtros de Fecha")
modo = st.sidebar.selectbox("Modo", ["Diario", "Semanal", "Mensual"])
fecha_seleccionada = st.sidebar.date_input("Selecciona fecha")

# Rango de fechas
if modo == "Diario":
    fecha_ini = pd.to_datetime(fecha_seleccionada)
    fecha_fin = fecha_ini
elif modo == "Semanal":
    fecha_ini = pd.to_datetime(fecha_seleccionada) - pd.to_timedelta(fecha_seleccionada.weekday(), unit='D')
    fecha_fin = fecha_ini + pd.Timedelta(days=5)
elif modo == "Mensual":
    fecha_ini = pd.to_datetime(fecha_seleccionada.replace(day=1))
    fecha_fin = fecha_ini + pd.offsets.MonthEnd(0)

# Filtros de datos
df_viajes = viajes[(viajes["FECHA"] >= fecha_ini) & (viajes["FECHA"] <= fecha_fin)]
formularios["Fecha_dia"] = formularios["Fecha de llenar"].dt.date
df_form = formularios[(formularios["Fecha_dia"] >= fecha_ini.date()) & (formularios["Fecha_dia"] <= fecha_fin.date())]

# Tabs
tab1, tab2, tab3 = st.tabs(["📌 Actividad", "📝 Formularios", "📍 Zonas & Alertas"])

with tab1:
    st.header("📌 Resumen de Actividad")
    st.markdown(f"📆 Desde: `{fecha_ini.date()}` hasta `{fecha_fin.date()}`")
    total_km = df_viajes["Distancia recorrida total km"].sum()
    total_viajes = len(df_viajes)
    total_tiempo = pd.to_timedelta(df_viajes["Tiempo de viaje"].astype(str), errors='coerce').sum()
    st.metric("🚗 KM recorridos", f"{total_km:.2f} KM")
    st.metric("🧭 Total de viajes", total_viajes)
    st.metric("⏱️ Tiempo total viaje", str(total_tiempo))

    if not df_viajes.empty:
        st.subheader("🔄 Viajes por día")
        viajes_dia = df_viajes.groupby("FECHA").size().reset_index(name="Cantidad")
        fig = px.bar(viajes_dia, x="FECHA", y="Cantidad", title="Desplazamientos por Día")
        st.plotly_chart(fig)

with tab2:
    st.header("📝 Formularios llenados")
    if df_form.empty:
        st.warning("No hay formularios en este período.")
    else:
        st.success(f"✅ Formularios registrados: {len(df_form)}")
        df_tipo = df_form['Tipo'].value_counts().reset_index()
        df_tipo.columns = ['Tipo', 'Cantidad']
        fig = px.pie(df_tipo, values='Cantidad', names='Tipo', title='Tipos de formularios llenados')
        st.plotly_chart(fig)

        st.subheader("📊 Formularios por empleado")
        form_emp = df_form.groupby('Empleado').size().reset_index(name="Cantidad")
        fig_emp = px.bar(form_emp, x='Empleado', y='Cantidad', title="Formularios llenados por empleado")
        st.plotly_chart(fig_emp)

        st.subheader("🔍 Detalle por cliente")
        cliente_sel = st.selectbox("Selecciona cliente", df_form['Tarea'].dropna().unique())
        df_cliente = df_form[df_form['Tarea'] == cliente_sel]

        for _, row in df_cliente.iterrows():
            with st.expander(f"{row['Fecha de llenar'].date()} - {row['Nombre de formulario']}"):
                st.markdown(f"📍 **Dirección:** {row['Dirección de envío']}")
                st.markdown(f"👤 **Doctor/Clínica:** {row['¿Cuál es el nombre del Doctor/ la Clínica?']}")
                st.markdown(f"📋 **Actividad:** {row['¿Qué actividades realizaste?']}")
                st.markdown(f"🧍 **Visitado:** {row['¿A quién visitaste?']}")
                st.markdown(f"🗒️ **Notas:** {row['Notas adicionales sobre la visita']}")
                if pd.notna(row['Evidencia Fotográfica']):
                    # Asumiendo que 'Evidencia Fotográfica' contiene URLs completas a imágenes
                    try:
                        st.image(row['Evidencia Fotográfica'], caption="Evidencia Fotográfica")
                    except:
                        st.warning("No se pudo cargar la imagen")

        st.subheader("📈 Actividad por día")
        productividad = df_form.groupby('Fecha_dia').size().reset_index(name="Formularios")
        fig2 = px.bar(productividad, x='Fecha_dia', y='Formularios', title="Formularios por fecha")
        st.plotly_chart(fig2)

        st.subheader("📤 Exportar PDF")
        if st.button("📄 Exportar PDF de Formularios"):
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.cell(200, 10, txt="Resumen de formularios", ln=True, align='C')
            pdf.cell(200, 10, txt=f"Periodo: {fecha_ini.date()} al {fecha_fin.date()}", ln=True)
            for _, row in df_form.iterrows():
                texto = f"{row['Fecha de llenar'].date()} - {row['Tarea']} - {row['¿Qué actividades realizaste?']}"
                try:
                    pdf.multi_cell(0, 10, texto.encode('latin-1', 'ignore').decode('latin-1'))
                except:
                    pdf.multi_cell(0, 10, "⚠️ Error en contenido especial")
            
            # Guardar PDF en memoria
            pdf_output = BytesIO()
            pdf.output(pdf_output)
            pdf_bytes = pdf_output.getvalue()
            
            st.download_button(
                label="📄 Descargar PDF",
                data=pdf_bytes,
                file_name="resumen_formularios.pdf",
                mime="application/pdf"
            )

with tab3:
    st.header("📍 Zonas visitadas (Mapa + Mapa de Calor)")
    df_map = df_form.dropna(subset=["Latitud", "Longitud"])
    if df_map.empty:
        st.warning("No hay coordenadas disponibles.")
    else:
        mapa = folium.Map(location=[18.5, -69.9], zoom_start=10)
        for _, row in df_map.iterrows():
            folium.CircleMarker(
                location=[row['Latitud'], row['Longitud']],
                radius=6,
                popup=f"{row['Tarea']} - {row['Fecha de llenar'].date()}",
                color='blue',
                fill=True,
                fill_opacity=0.6
            ).add_to(mapa)
        heat_data = df_map[['Latitud', 'Longitud']].values.tolist()
        HeatMap(heat_data).add_to(mapa)
        st_folium(mapa, width=700, height=500)

        st.divider()
        st.subheader("🚨 Alertas automáticas")
        if modo == "Semanal":
            if total_km < 10:
                st.error("⚠️ Menos de 10 KM esta semana.")
            if len(df_form) < 2:
                st.warning("⚠️ Menos de 2 formularios esta semana.")

        st.subheader("🏆 Ranking semanal por formularios")
        ranking = df_form.groupby('Empleado').size().reset_index(name="Formularios").sort_values("Formularios", ascending=False)
        st.dataframe(ranking)
        fig3 = px.bar(ranking, x='Empleado', y='Formularios', title="Ranking semanal")
        st.plotly_chart(fig3)
        
# --- Recomendaciones ---
st.sidebar.header("Recomendaciones")
st.sidebar.markdown("""
- Implementar un sistema de notificaciones para alertas críticas.
- Añadir un módulo de seguimiento de objetivos mensuales.
- Integrar geocodificación previa para mostrar mapas de calor con latitud y longitud reales.
- Añadir alertas personalizadas por vendedor y productos visitados.
- Agregar análisis de tendencias mensuales para detectar patrones de mercado.
- Incorporar módulo de feedback para que vendedores reporten incidencias o sugerencias.
- Optimizar la UI para móvil con `streamlit-mobile` o diseño responsive.
- Automatizar envío de reportes semanales por correo a gerencia.
""")
