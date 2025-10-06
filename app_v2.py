# ==============================================================================
#     1. IMPORTACIONES Y CONFIGURACIÓN INICIAL
# ==============================================================================
import streamlit as st
import pandas as pd
import numpy as np
from scipy import stats
import io

# --- CONFIGURACIÓN DE LA PÁGINA DE STREAMLIT ---
st.set_page_config(
    page_title="Análisis de Métricas IR & LOI",
    page_icon="📊",
    layout="wide"
)

# --- FUNCIÓN PARA APLICAR ESTILOS CSS PERSONALIZADOS ---
def aplicar_estilos_personalizados():
    estilos = """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap');
        @import url('https://fonts.googleapis.com/css2?family=Hind:wght@400;500;600&display=swap');
        :root {
            --atlantia-violet: #6546C3; --atlantia-purple: #AA49CA; --atlantia-lemon: #77C014;
            --atlantia-turquoise: #04D1CD; --atlantia-white: #FFFFFF; --atlantia-green: #23B776;
            --atlantia-yellow: #FFB73B; --atlantia-orange: #FF9231; --atlantia-red: #E61252;
            --atlantia-black: #000000;
        }
        html, body, [class*="st-"] { font-family: 'Hind', sans-serif; font-size: 12pt; color: var(--atlantia-black); }
        h1, h2, h3 { font-family: 'Poppins', sans-serif !important; font-weight: 700 !important; color: var(--atlantia-violet) !important; }
        h1 { font-size: 24pt !important; } h2 { font-size: 20pt !important; } h3 { font-size: 18pt !important; }
        .metric-container { background: white; border: 2px solid var(--atlantia-turquoise); border-radius: 10px; padding: 1.5rem 1rem 1rem 1rem; text-align: center; box-shadow: 0 2px 10px rgba(4, 209, 205, 0.1); }
        .metric-label { font-family: 'Hind', sans-serif !important; font-weight: 500 !important; font-size: 14pt !important; color: var(--atlantia-violet) !important; }
        .metric-value { font-family: 'Poppins', sans-serif !important; font-weight: 700 !important; font-size: 22pt !important; color: var(--atlantia-black) !important; margin-top: 0.5rem; }
        .stButton > button { font-family: 'Hind', sans-serif !important; font-weight: 600 !important; font-size: 12pt !important; background: linear-gradient(135deg, var(--atlantia-violet) 0%, var(--atlantia-purple) 100%) !important; border: none !important; border-radius: 8px !important; box-shadow: 0 4px 15px rgba(101, 70, 195, 0.3) !important; transition: all 0.3s ease !important; }
        .stButton > button:hover { transform: translateY(-2px) !important; box-shadow: 0 6px 20px rgba(101, 70, 195, 0.4) !important; }
        .stAlert { border-radius: 10px !important; }
        .stAlert[data-baseweb="notification-positive"] { background-color: rgba(35, 183, 118, 0.2) !important; color: var(--atlantia-green) !important;}
        .stAlert[data-baseweb="notification-negative"] { background-color: rgba(230, 18, 82, 0.2) !important; color: var(--atlantia-red) !important;}
        .stAlert[data-baseweb="notification-info"] { background-color: rgba(4, 209, 205, 0.2) !important; color: var(--atlantia-turquoise) !important;}
        .stFileUploader > div > div { border: 2px dashed var(--atlantia-violet) !important; border-radius: 10px !important; background-color: rgba(101, 70, 195, 0.05) !important; }
        .st-expander > summary { font-family: 'Hind', sans-serif !important; font-weight: 600 !important; font-size: 14pt !important; color: var(--atlantia-violet) !important; }
        .st-expander > summary:hover { color: var(--atlantia-purple) !important; }
    </style>
    """
    st.markdown(estilos, unsafe_allow_html=True)


# --- CONSTANTES Y CONFIGURACIONES DEL ANÁLISIS ---
COL_FILTRO = 'Filtro'
COL_TIEMPO_BASE = 'Tiempo total'
POSICIONES_HOJAS_ESTANDAR = {"efectivas": 0, "todas1": 1, "todas2": 2, "completadas2": 3}
POSICIONES_HOJAS_EXCEPCION = {"efectivas": 0, "todas1": 1, "completadas2": 2}

# --- FUNCIÓN PARA CREAR LA TABLA DE STATUS CON HTML ---
def crear_tabla_status_html(df):
    header_style = "text-align: left; padding: 8px; border-bottom: 2px solid var(--atlantia-violet); color: var(--atlantia-violet);"
    cell_style = "padding: 8px; border-bottom: 1px solid #eee;"
    html = f"""<table style="width:100%; border-collapse: collapse; font-family: 'Hind', sans-serif;"><thead><tr><th style="{header_style}">Status</th><th style="{header_style}">Conteo</th><th style="{header_style}">Porcentaje (%)</th></tr></thead><tbody>"""
    for _, row in df.iterrows():
        status, conteo, percent = row['Status'], row['Conteo'], row['Porcentaje']
        progress_bar_html = f"""<div style="display: flex; align-items: center; width: 100%;"><div style="width: {percent}%; background-color: #6546C3; height: 20px; border-radius: 5px;"></div><span style="padding-left: 8px; white-space: nowrap;">{percent:.2f}%</span></div>"""
        html += f"""<tr><td style="{cell_style}">{status}</td><td style="{cell_style}">{conteo}</td><td style="{cell_style} width: 50%;">{progress_bar_html}</td></tr>"""
    html += "</tbody></table>"
    return html

# ==============================================================================
#     2. FUNCIÓN PARA PROCESAR LOS DATOS CARGADOS
# ==============================================================================
def procesar_datos_excel(archivo_excel, num_hojas):
    try:
        if num_hojas >= 4:
            posiciones_a_leer = POSICIONES_HOJAS_ESTANDAR
            sheet_indices = list(posiciones_a_leer.values())
            data_sheets = pd.read_excel(archivo_excel, sheet_name=sheet_indices)
            df_efectivas, df_todas1, df_todas2, df_completadas2 = (data_sheets[pos] for pos in sheet_indices)
        else:
            posiciones_a_leer = POSICIONES_HOJAS_EXCEPCION
            sheet_indices = list(posiciones_a_leer.values())
            data_sheets = pd.read_excel(archivo_excel, sheet_name=sheet_indices)
            df_efectivas, df_todas1, df_completadas2 = (data_sheets[pos] for pos in sheet_indices)
            df_todas2 = None
    except Exception as e: return {"error": f"Ocurrió un error al leer las hojas del archivo: {e}"}
    
    dfs_a_limpiar = [df for df in [df_efectivas, df_todas1, df_todas2, df_completadas2] if df is not None]
    for df in dfs_a_limpiar: df.columns = df.columns.str.strip()

    try:
        # --- CAMBIO: Lógica de estandarización de columnas de ID ---
        id_unificado = 'id_unificado'
        if num_hojas >= 4:
            # Caso estándar: [auth]
            col_auth_original = '[auth]'
            df_efectivas.rename(columns={col_auth_original: id_unificado}, inplace=True)
            df_todas1.rename(columns={col_auth_original: id_unificado}, inplace=True)
            df_todas2.rename(columns={col_auth_original: id_unificado}, inplace=True)
            df_completadas2.rename(columns={col_auth_original: id_unificado}, inplace=True)
            
            df_tiempos_p2 = df_todas2[[id_unificado, COL_TIEMPO_BASE]].drop_duplicates(subset=[id_unificado], keep='first')
            df_final = pd.merge(df_todas1, df_tiempos_p2, on=id_unificado, how='left', suffixes=('_p1', '_p2'))
            df_final['duracion_total_seg'] = df_final[f"{COL_TIEMPO_BASE}_p1"].fillna(0) + df_final[f"{COL_TIEMPO_BASE}_p2"].fillna(0)
        else:
            # Caso excepcional: 'id' y 'ID de respuesta'
            df_efectivas.rename(columns={'id': id_unificado}, inplace=True)
            df_todas1.rename(columns={'ID de respuesta': id_unificado}, inplace=True)
            df_completadas2.rename(columns={'ID de respuesta': id_unificado}, inplace=True)
            
            df_final = df_todas1.copy()
            df_final['duracion_total_seg'] = df_final[COL_TIEMPO_BASE].fillna(0)

        # --- A partir de aquí, el código usa 'id_unificado' y funciona para ambos casos ---
        auth_efectivos = set(df_efectivas[id_unificado])
        auth_solo_completadas = set(df_completadas2[id_unificado])
        
        conditions = [
            df_final[id_unificado].isin(auth_efectivos),
            df_final[COL_FILTRO].notna(),
            (df_final[COL_FILTRO].isna()) & (df_final[id_unificado].isin(auth_solo_completadas))
        ]
        choices = ['Completada', 'Filtrada', 'Descartada']
        df_final['status_final'] = np.select(conditions, choices, default='Incompleta')
        
        resumen_status = df_final['status_final'].value_counts().reset_index()
        resumen_status.columns = ['Status', 'Conteo']
        total_encuestas = df_final.shape[0]
        if total_encuestas > 0:
            resumen_status['Porcentaje'] = (resumen_status['Conteo'] / total_encuestas) * 100
        else:
            resumen_status['Porcentaje'] = 0
            
        ir = 0
        if 'Completada' in resumen_status['Status'].values:
            conteo_completadas = resumen_status.loc[resumen_status['Status'] == 'Completada', 'Conteo'].iloc[0]
            if len(df_final) > 0: ir = conteo_completadas / len(df_final)
            
        media_acotada_minutos = 0
        df_completadas = df_final[df_final['status_final'] == 'Completada'].copy()
        if not df_completadas.empty:
            duraciones_validas = df_completadas['duracion_total_seg'].dropna()
            if not duraciones_validas.empty:
                media_acotada_segundos = stats.trim_mean(duraciones_validas, 0.05)
                media_acotada_minutos = media_acotada_segundos / 60
                
        return {"error": None, "df_final": df_final, "resumen_status": resumen_status, "ir": ir, "loi_minutos": media_acotada_minutos}
        
    except KeyError as e:
        return {"error": f"No se encontró una columna de identificación esperada: {e}. Verifica que el archivo Excel tenga '[auth]' o 'id'/'ID de respuesta' según corresponda."}
    except Exception as e:
        return {"error": f"Ocurrió un error durante el procesamiento: {e}"}

# ==============================================================================
#     3. FUNCIÓN PARA CONVERTIR DATAFRAME A EXCEL EN MEMORIA
# ==============================================================================
@st.cache_data
def convertir_a_excel(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_dict_copy = {name: df.copy() for name, df in df_dict.items()}
        if 'Resumen_Estatus' in df_dict_copy:
            resumen_df = df_dict_copy['Resumen_Estatus']
            if 'Porcentaje' in resumen_df.columns:
                resumen_df['Porcentaje'] = resumen_df['Porcentaje'] / 100.0
        for sheet_name, df in df_dict_copy.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            if sheet_name == 'KPIs':
                percent_format = workbook.add_format({'num_format': '0.00%', 'bold': True})
                float_format = workbook.add_format({'num_format': '0.00', 'bold': True})
                worksheet.write('B2', df.iloc[0, 1], percent_format)
                worksheet.write('B3', df.iloc[1, 1], float_format)
                worksheet.set_column('A:A', 40)
                worksheet.set_column('B:B', 15)
            if sheet_name == 'Resumen_Estatus':
                percent_format = workbook.add_format({'num_format': '0.00%'})
                if 'Porcentaje' in df.columns:
                    col_idx = df.columns.get_loc('Porcentaje')
                    worksheet.set_column(col_idx, col_idx, 15, percent_format)
    output.seek(0)
    return output.getvalue()

# ==============================================================================
#     4. INTERFAZ DE USUARIO CON STREAMLIT
# ==============================================================================
aplicar_estilos_personalizados()

st.markdown("""<div style="text-align: center;"><h1>📊 Análisis de Métricas IR & LOI</h1><p style="color: var(--atlantia-violet); font-weight: 500;">Una herramienta automatizada por Atlantia</p></div><hr style="border-top: 2px solid var(--atlantia-purple); margin-bottom: 2rem;">""", unsafe_allow_html=True)
st.markdown("""### Instrucciones:
Asegúrate de que tu archivo `.xlsx` siga una de las dos estructuras:
-   **Estándar (4 Hojas):**
    1.  Encuestas efectivas (numérica)
    2.  Todas las Encuestas (Parte 1)
    3.  Todas las Encuestas (Parte 2)
    4.  Encuestas Completadas (Parte 2)
-   **Excepción (3 Hojas):**
    1.  Encuestas efectivas (numérica)
    2.  Todas las Encuestas (Parte 1)
    3.  Encuestas Completadas (Parte 2)""")

uploaded_file = st.file_uploader("Carga tu archivo Excel de métricas", type=['xlsx'])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        num_hojas = len(xls.sheet_names)
        if num_hojas < 3: st.error(f"❌ **Error:** El archivo cargado solo tiene {num_hojas} hoja(s). Se requiere un mínimo de 3.")
        else:
            st.info(f"Archivo detectado con {num_hojas} hoja(s). Procesando según corresponda...")
            with st.spinner("Procesando archivo... por favor espera."):
                resultados = procesar_datos_excel(uploaded_file, num_hojas)
            if resultados["error"]: st.error(f"❌ **Error:** {resultados['error']}")
            else:
                st.success("✅ ¡Análisis completado con éxito!")
                df_final, resumen_status, ir_calculado, loi_calculado = (
                    resultados["df_final"], resultados["resumen_status"],
                    resultados["ir"], resultados["loi_minutos"]
                )
                st.header("Resultados del Análisis")
                col1, col2, col3 = st.columns(3)
                with col1: st.markdown(f'<div class="metric-container"><div class="metric-label">Tasa de Incidencia (IR)</div><div class="metric-value">{ir_calculado:.2%}</div></div>', unsafe_allow_html=True)
                with col2: st.markdown(f'<div class="metric-container"><div class="metric-label">Duración de Entrevista (LOI)</div><div class="metric-value">{loi_calculado:.2f} min</div></div>', unsafe_allow_html=True)
                with col3: st.markdown(f'<div class="metric-container"><div class="metric-label">Total de Registros Analizados</div><div class="metric-value">{len(df_final)}</div></div>', unsafe_allow_html=True)
                
                st.subheader("Resumen por status del total de encuestas")
                tabla_html = crear_tabla_status_html(resumen_status)
                st.markdown(tabla_html, unsafe_allow_html=True)

                st.header("Descarga de Resultados")
                kpi_data = {
                    'Métrica': [
                        'Tasa de Incidencia (IR)',
                        'Duración de Entrevista (LOI en minutos)',
                        'Total de Registros Analizados'
                    ],
                    'Valor': [ir_calculado, loi_calculado, len(df_final)]
                }
                df_kpis = pd.DataFrame(kpi_data)
                data_to_download = {
                    "KPIs": df_kpis,
                    "Resumen_Estatus": resumen_status,
                    "Base_Procesada_Completa": df_final
                }
                excel_bytes = convertir_a_excel(data_to_download)
                st.download_button(
                    label="📥 Descargar Análisis en Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Analisis_IR_LOI_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                with st.expander("Ver la base de datos completa procesada"):
                    st.dataframe(df_final)
    except Exception as e:
        st.error(f"❌ **Error:** No se pudo leer el archivo Excel. Puede que esté dañado. Detalle: {e}")
else:
    st.info("Esperando a que cargues un archivo Excel para comenzar el análisis.")

st.markdown("""<hr style="border-top: 1px solid var(--atlantia-turquoise); margin-top: 3rem;"><div style="text-align: center; color: var(--atlantia-violet); font-weight: 500; padding: 1rem 0;">Powered by Atlantia | Octubre 2025</div>""", unsafe_allow_html=True)