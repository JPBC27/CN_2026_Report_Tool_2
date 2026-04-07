import streamlit as st
import os
import pandas as pd
import io
from datetime import datetime
from reporte_cn_2026 import (
    normalizar_texto, 
    procesar_archivo_dotacion, 
    procesar_archivo_censo, 
    procesar_archivo_adicional, 
    procesar_archivo_cesados,
    aplicar_purga_cesados,
    enriquecer_final,
    cargar_config_cursos,
    guardar_config_cursos,
    procesar_ucenco,
    procesar_capacitacion_comun,
    consolidar_capacitaciones
)

# --- Interfaz de Streamlit ---
st.set_page_config(
    page_title="Consolidación CN 2026 (Versión Full)",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo Premium (CSS Inyectado)
st.markdown("""
    <style>
    .main {
        background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #020617 100%);
        color: white;
    }
    .stButton>button {
        background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%);
        border: none;
        color: white;
        padding: 0.75rem 1.75rem;
        border-radius: 14px;
        font-weight: 700;
        letter-spacing: 0.025em;
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
    }
    .stButton>button:hover {
        transform: translateY(-3px) scale(1.02);
        box-shadow: 0 20px 25px -5px rgba(99, 102, 241, 0.3), 0 10px 10px -5px rgba(99, 102, 241, 0.2);
    }
    h1 {
        background: linear-gradient(to right, #818cf8, #c084fc, #f472b6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 900 !important;
        letter-spacing: -0.025em;
    }
    .stSidebar {
        background: rgba(255, 255, 255, 0.03) !important;
        backdrop-filter: blur(20px);
        border-right: 1px solid rgba(255, 255, 255, 0.1);
    }
    /* Glassmorphism card effect */
    .glass-card {
        background: rgba(255, 255, 255, 0.05);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 20px;
        padding: 24px;
        margin: 10px 0;
    }
    /* Metric styling */
    [data-testid="stMetricValue"] {
        font-size: 2.2rem !important;
        font-weight: 800 !important;
        background: linear-gradient(to bottom, #ffffff, #94a3b8);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    </style>
""", unsafe_allow_html=True)

st.title('🚀 Consolidación y Reporte CN 2026')
st.write("---")

st.markdown("""
### Bienvenido a la herramienta centralizada de RR.HH.
Esta plataforma permite consolidar la **Dotación General** con los nuevos ingresos de **CENSO_PERU**, aplicando filtros predictivos y enriqueciendo los datos mediante cruces automáticos de información estratégica.
""")

# Sidebar: Configuración de Archivos
st.sidebar.image("https://img.icons8.com/isometric-line/100/6366f1/data-configuration.png", width=80)
st.sidebar.header("📥 Archivos de Entrada")

uploaded_files = st.sidebar.file_uploader(
    "Selecciona uno o más archivos (.xlsx, .xls, .xlsm)", 
    type=["xlsx", "xls", "xlsm", "xlsb"], 
    accept_multiple_files=True
)

# Selector de Acción
action = st.selectbox(
    "Selecciona el tipo de operación:",
    ["Reporte de Cursos Normativos", "Análisis de Datos (Próximamente)", "Configuración de Mapeo"]
)

st.divider()

if action == "Reporte de Cursos Normativos":
    st.subheader("Configuración: Reporte de Cursos Normativos")
    st.write("""
    **Archivos requeridos en el panel lateral:**
    1.  `DOTACIÓN GENERAL`
    2.  `CENSO_PERU...` (uno o varios)
    3.  `Información Adicional`
    4.  `Maestro Cesados`
    """)
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        if st.button('🚀 Generar Reporte Full', type="primary", use_container_width=True):
            if uploaded_files:
                with st.status("🚀 Iniciando Consolidación de Archivos...", expanded=True) as status:
                    # 1. Identificar archivos
                    file_dot = None
                    list_censo = []
                    file_info = None
                    file_cesados = None
                    list_ucenco = []
                    list_campus = []
                    
                    for f in uploaded_files:
                        n_norm = normalizar_texto(f.name)
                        n_raw = f.name.upper()
                        if 'DOTACION GENERAL' in n_norm: file_dot = f
                        elif 'CENSO' in n_norm: list_censo.append(f)
                        elif 'INFORMACION ADICIONAL' in n_norm: file_info = f
                        elif 'CESADOS' in n_norm: file_cesados = f
                        elif n_raw.startswith('UCENCO_'): list_ucenco.append(f)
                        elif 'UCENCO' in n_norm:
                            # Legacy or direct UCENCO name match
                            list_ucenco.append(f)
                        elif f.name.startswith('Campus_'): list_campus.append(f)

                    if not file_dot:
                        st.error("⚠️ Archivo 'DOTACIÓN GENERAL' es obligatorio.")
                    else:
                        # 2. Procesar Dotación (Independiente)
                        df_dot, max_alta, subdir_map = procesar_archivo_dotacion(file_dot)
                        
                        if df_dot is not None:
                            # 3. Procesar Censo (Archivo por archivo)
                            censo_dfs = []
                            for f_censo in list_censo:
                                df_c_ind = procesar_archivo_censo(f_censo, max_alta, subdir_map)
                                if not df_c_ind.empty:
                                    censo_dfs.append(df_c_ind)

                            # 4. Consolidar Base
                            if censo_dfs:
                                df_censo_full = pd.concat(censo_dfs, ignore_index=True)
                                sap_exists = set(df_dot['Nº pers.'].astype(str).str.strip())
                                df_censo_full = df_censo_full[~df_censo_full['Nº pers.'].isin(sap_exists)]
                                st.write(f"✅ Nuevos ingresos descubiertos: **{len(df_censo_full)}**")
                                df_final = pd.concat([df_dot, df_censo_full], ignore_index=True)
                            else:
                                df_final = df_dot.copy()

                            # 5. Enriquecer con Info Adicional (Independiente)
                            info_adicional = {}
                            if file_info:
                                info_adicional = procesar_archivo_adicional(file_info)
                            
                            df_final = enriquecer_final(df_final, info_adicional)
                            
                            # 6. Purga de Cesados (Si aplica)
                            eliminados_count = 0
                            if file_cesados:
                                df_cesados_data = procesar_archivo_cesados(file_cesados)
                                # Recibimos consolidado limpio Y lista de eliminados
                                df_final, df_eliminados = aplicar_purga_cesados(df_final, df_cesados_data)
                                st.session_state['df_cesados_final'] = df_eliminados
                                eliminados_count = len(df_eliminados)
                            
                            # 7. Procesar Capacitaciones (Si existen)
                            df_u_data = None
                            df_c_data = None
                            
                            if list_ucenco:
                                list_u_dyn = info_adicional.get('lista_ucenco')
                                ucenco_dfs = [procesar_ucenco(f, cursos_validos_custom=list_u_dyn) for f in list_ucenco]
                                df_u_data = pd.concat(ucenco_dfs, ignore_index=True) if ucenco_dfs else None
                            
                            if list_campus:
                                list_c_dyn = info_adicional.get('lista_campus')
                                campus_dfs = [procesar_capacitacion_comun(f, f.name.replace('Campus_', '').replace('.xlsx', '').replace('.xls', ''), cursos_validos_custom=list_c_dyn) for f in list_campus]
                                df_c_data = pd.concat(campus_dfs, ignore_index=True) if campus_dfs else None
                                
                            if df_u_data is not None or df_c_data is not None:
                                df_final = consolidar_capacitaciones(df_final, df_u_data, df_c_data, info_adicional)

                            # Guardar en Session State
                            st.session_state['df_consolidado'] = df_final
                            
                            # Reporte Estadístico Final
                            st.divider()
                            st.subheader("📊 Resumen del Proceso")
                            
                            # Calcular Cumplimiento General si existen las columnas
                            promedios = ""
                            if '% Cumplimiento' in df_final.columns:
                                cumple_gral = df_final['% Cumplimiento'].mean()
                                promedios = f"{cumple_gral:.1f}%"
                            
                            m1, m2, m3, m4 = st.columns(4)
                            m1.metric("Base Final", f"{len(df_final):,}")
                            m2.metric("Nuevos (Censo)", f"{len(df_censo_full) if censo_dfs else 0:,}")
                            m3.metric("Cesados", f"{eliminados_count:,}")
                            if promedios:
                                m4.metric("Cumplimiento Gral.", promedios)
                            else:
                                m4.metric("Dotación Base", f"{len(df_dot):,}")
                            
                            # Preparar Descarga con Estilos
                            with st.spinner("🎨 Aplicando diseño corporativo al reporte..."):
                                temp_file = "temp_reporte.xlsx"
                                df_final.to_excel(temp_file, index=False, sheet_name='Reporte_Consolidado')
                                
                                # Aplicar Estilos OpenPyXL (incluyendo Dashboard)
                                from reporte_cn_2026 import aplicar_estilos_custom
                                if aplicar_estilos_custom(temp_file, info_adicional, df_final):
                                    with open(temp_file, "rb") as f:
                                        st.session_state['output_excel'] = f.read()
                                    if os.path.exists(temp_file):
                                        os.remove(temp_file)
                                else:
                                    # Fallback: si falla el estilo, devolver el buffer simple
                                    output = io.BytesIO()
                                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                        df_final.to_excel(writer, index=False, sheet_name='Reporte_Consolidado')
                                    st.session_state['output_excel'] = output.getvalue()

                            status.update(label="✅ Consolidación Finalizada", state="complete", expanded=False)
                            st.toast("✅ Procesamiento completado con éxito", icon="🎉")
            else:
                st.error("⚠️ Por favor, sube los archivos requeridos.")
    
    with col2:
        if uploaded_files:
            st.success(f"Archivos listos: {len(uploaded_files)} detectados.")
        else:
            st.warning("Sin archivos cargados.")

    if 'df_consolidado' in st.session_state:
        st.divider()
        st.subheader("📥 Resultados")
        
        c_down1, c_down2, c_blank = st.columns([1, 1, 1])
        with c_down1:
            st.download_button(
                label="📥 Descargar Reporte Final (XLSX)", 
                data=st.session_state['output_excel'], 
                file_name=f"Reporte_Consolidado_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with c_down2:
            if 'df_cesados_final' in st.session_state and not st.session_state['df_cesados_final'].empty:
                csv_cesados = st.session_state['df_cesados_final'].to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Descargar Eliminados (CSV)",
                    data=csv_cesados,
                    file_name=f"Colaboradores_Eliminados_{datetime.now().strftime('%d_%m_%Y')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            

elif action == "Análisis de Datos (Próximamente)":
    st.subheader("📊 Análisis Visual de Dotación y Capacitación")
    
    if 'df_consolidado' not in st.session_state:
        st.info("💡 Primero genera el reporte en la pestaña anterior para ver el análisis.")
    else:
        df = st.session_state['df_consolidado'].copy()
        
        # Métricas Clave
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Colaboradores", f"{len(df):,}")
        
        # Intentar detectar nuevos ingresos por año actual en la columna 'Alta'
        current_year = str(datetime.now().year)
        nuevos = len(df[df['Alta'].astype(str).str.contains(current_year, na=False)])
        m2.metric("Nuevos (Censo)", nuevos)
        
        if '% Cumplimiento' in df.columns:
            m3.metric("Cumplimiento Promedio", f"{df['% Cumplimiento'].mean():.1f}%")
            completos = len(df[df['ESTADO FINAL'] == "COMPLETADO"])
            m4.metric("Colab. Completados", f"{completos:,}")
        else:
            m3.metric("Banderas Únicas", df['Bandera'].nunique() if 'Bandera' in df.columns else 0)
            m4.metric("Secciones", df['Sección'].nunique() if 'Sección' in df.columns else 0)
        
        st.divider()
        
        # --- FILA 1: GEOGRAFÍA Y ESTRUCTURA ---
        st.markdown("### 🏢 Estructura Organizacional")
        c1, c2 = st.columns(2)
        
        with c1:
            st.markdown("#### 🚩 Distribución por Bandera")
            if 'Bandera' in df.columns and not df['Bandera'].empty:
                bandera_counts = df['Bandera'].value_counts()
                st.bar_chart(bandera_counts)
        
        with c2:
            st.markdown("#### 📍 Top 10 Ciudades")
            if 'Ciudad' in df.columns and not df['Ciudad'].empty:
                city_counts = df['Ciudad'].value_counts().head(10)
                st.bar_chart(city_counts)
        
        # --- FILA 2: CAPACITACIÓN ---
        if '% Cumplimiento' in df.columns:
            st.divider()
            st.markdown("### 🎓 Indicadores de Capacitación")
            
            c3, c4 = st.columns(2)
            
            with c3:
                st.markdown("#### 👤 Cumplimiento por Gerente Zonal")
                if 'Gte. Zonal' in df.columns:
                    # Agrupar y promediar
                    gz_comp = df.groupby('Gte. Zonal')['% Cumplimiento'].mean().sort_values(ascending=False)
                    st.bar_chart(gz_comp)
                else:
                    st.warning("Columna 'Gte. Zonal' no disponible para análisis.")
            
            with c4:
                st.markdown("#### 📈 Estado Final de Colaboradores")
                if 'ESTADO FINAL' in df.columns:
                    est_counts = df['ESTADO FINAL'].value_counts()
                    st.bar_chart(est_counts)
            
            # Detalle por Sección
            st.markdown("#### ⚠️ Top 5 Secciones con Menor Cumplimiento")
            if 'Sección' in df.columns:
                sec_comp = df.groupby('Sección')['% Cumplimiento'].mean().sort_values().head(5)
                st.dataframe(sec_comp.reset_index(), use_container_width=True)

elif action == "Configuración de Mapeo":
    st.subheader("⚙️ Estructura de Datos (Cabeceras Fijas)")
    st.info("El sistema está configurado para procesar archivos con cabeceras estandarizadas.")
    
    with st.expander("📋 Ver Cabeceras Requeridas"):
        st.markdown("""
        **Dotación General:**
        - `Nº pers.`, `Nombre del empleado o candidat`, `Alta`, `SubPer`, `Subdivisión de personal`, `Función`, `Fecha`, `ID Number`, `Bandera`, `Clasificación`, `PCD`, `Dependencia`

        **Censo Perú:**
        - `Cod. Sap`, `Nombre`, `Cd Ubicación`, `Ubicación`, `Posición`, `Puesto`, `Departamento`, `Área de Trabajo`, `Doc ID`, `Fecha Ingreso Planilla`
        """)
    
    st.success("La detección automática ha sido reemplazada por mapeo directo de alta precisión.")

    st.divider()
    st.subheader("🎓 Gestión de Cursos (Configuración)")
    st.info("💡 **Nota:** La lista de cursos ahora es **dinámica**. Se extrae automáticamente de la hoja 'Lista de Cursos' del archivo 'Información Adicional'. Los valores a continuación se usarán solo como respaldo si no se encuentra dicha hoja.")
    
    config = cargar_config_cursos()
    
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        st.markdown("#### Cursos UCENCO")
        ucenco_text = st.text_area("Un curso por línea (exacto como en Excel):", value="\n".join(config['ucenco']), height=150)
        config['ucenco'] = [c.strip() for c in ucenco_text.split("\n") if c.strip()]
        
    with col_c2:
        st.markdown("#### Cursos Transversales / SSO")
        trans_text = st.text_area("Un curso por línea (exacto como en Excel):", value="\n".join(config['transversales_sso']), height=150)
        config['transversales_sso'] = [c.strip() for c in trans_text.split("\n") if c.strip()]
        
    if st.button("💾 Guardar Configuración de Cursos"):
        guardar_config_cursos(config)
        st.success("✅ Configuración de cursos guardada correctamente.")
        st.toast("Configuración actualizada")

st.sidebar.divider()
st.sidebar.caption("v1.1.0 | Dashboard & Refactor")
