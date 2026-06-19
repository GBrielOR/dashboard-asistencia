import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl

st.set_page_config(page_title="Control de Asistencias - Unidad Educativa Caranqui", layout="wide")

st.title("Dashboard de Asistencias Escolares")
st.markdown("### Unidad Educativa Caranqui - Periodo Lectivo 2025 - 2026")

ARCHIVO_EXCEL = "Asistencia_Escolar_CORREGIDA_FINAL.xlsx"

@st.cache_data
def procesar_todo_el_excel(ruta_archivo):
    xl = pd.ExcelFile(ruta_archivo)
    hojas = xl.sheet_names
    
    lista_registros = []
    meses_dict = {
        "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
        "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12
    }
    
    progreso_texto = st.empty()
    
    for hoja in hojas:
        hoja_limpia = hoja.strip()
        if hoja_limpia.lower() in ["inicio", ""] or len(hoja_limpia) < 2:
            continue
            
        progreso_texto.text(f"Cargando datos: {hoja_limpia}")
            
        año = "2026"
        if "2025" in hoja_limpia:
            año = "2025"
        elif "2026" in hoja_limpia:
            año = "2026"
            
        curso = hoja_limpia.replace("2025", "").replace("2026", "").replace("-", "").strip()
        if curso == "": 
            curso = hoja_limpia
            
        try:
            df_hoja = pd.read_excel(ruta_archivo, sheet_name=hoja, header=None)
        except Exception:
            continue
            
        mes_actual = None
        
        for idx, row in df_hoja.iterrows():
            if len(row) < 2:
                continue
                
            valor_celda_1 = str(row.iloc[0]).strip().upper()
            valor_celda_2 = str(row.iloc[1]).strip().upper()
            
            if "ASISTENCIA" in valor_celda_1 or "ASISTENCIA" in valor_celda_2:
                continue
                
            if valor_celda_1 in meses_dict:
                mes_actual = valor_celda_1
                continue
                
            if "ESTUDIANTE" in valor_celda_2 or "N°" in valor_celda_1:
                continue
                
            estudiante = row.iloc[1]
            if pd.notna(estudiante) and str(estudiante).strip() != "":
                nom_estudiante = str(estudiante).strip().upper()
                if "TOTAL" in nom_estudiante or "ESTUDIANTE" in nom_estudiante or nom_estudiante in ["0", "0.0"]:
                    continue
                    
                for dia in range(1, 32):
                    col_idx = dia + 1
                    if col_idx < len(row):
                        estado = row.iloc[col_idx]
                        
                        if pd.isna(estado) or str(estado).strip() in ["", "0", "0.0"]:
                            estado_final = "P"
                        else:
                            estado_final = str(estado).strip().upper()
                            
                        if estado_final in ['P', 'A', 'T', 'J']:
                            lista_registros.append({
                                "Año": año,
                                "Curso": curso,
                                "Estudiante": str(estudiante).strip(),
                                "Mes": mes_actual if mes_actual else "NO ESPECIFICADO",
                                "Dia": dia,
                                "Estado": estado_final
                            })
                            
    progreso_texto.empty()
    
    if len(lista_registros) == 0:
        return pd.DataFrame()
        
    df_final = pd.DataFrame(lista_registros)
    mapa_estados = {'P': 'Presente', 'A': 'Ausente (Falta)', 'T': 'Tarde (Atraso)', 'J': 'Justificado'}
    df_final['Estado_Nombre'] = df_final['Estado'].map(mapa_estados)
    return df_final

try:
    df_asistencias = procesar_todo_el_excel(ARCHIVO_EXCEL)
    
    if df_asistencias.empty:
        st.error("No se pudieron extraer registros. Compruebe la estructura de las hojas del archivo Excel.")
    else:
        st.sidebar.header("Filtros de Control")
        
        lista_años = sorted(df_asistencias['Año'].unique())
        filtro_año = st.sidebar.selectbox("Seleccione el Año Lectivo:", options=lista_años)
        df_filtrado_año = df_asistencias[df_asistencias['Año'] == filtro_año]
        
        lista_cursos = sorted(df_filtrado_año['Curso'].unique())
        filtro_curso = st.sidebar.selectbox("Seleccione el Curso / Grado:", options=["Todos los Cursos"] + lista_cursos)
        
        if filtro_curso != "Todos los Cursos":
            df_final_display = df_filtrado_año[df_filtrado_año['Curso'] == filtro_curso]
        else:
            df_final_display = df_filtrado_año

        total_registros = len(df_final_display)
        if total_registros > 0:
            asistencias_ok = len(df_final_display[df_final_display['Estado'].isin(['P', 'T'])])
            porcentaje_asistencia = (asistencias_ok / total_registros) * 100
            total_faltas = len(df_final_display[df_final_display['Estado'] == 'A'])
            total_atrasos = len(df_final_display[df_final_display['Estado'] == 'T'])
            
            kpi1, kpi2, kpi3 = st.columns(3)
            kpi1.metric(label="Tasa de Asistencia Promedio", value=f"{porcentaje_asistencia:.1f}%")
            kpi2.metric(label="Total Inasistencias (Faltas)", value=f"{total_faltas:,}")
            kpi3.metric(label="Total de Atrasos Registrados", value=f"{total_atrasos:,}")
            
            st.markdown("---")
            
            col_izq, col_der = st.columns(2)
            
            with col_izq:
                st.subheader("Distribución General de Asistencias")
                fig_pastel = px.pie(df_final_display, names='Estado_Nombre', 
                                    color='Estado_Nombre',
                                    color_discrete_map={
                                        'Presente': '#2ecc71',
                                        'Ausente (Falta)': '#e74c3c',
                                        'Tarde (Atraso)': '#f1c40f',
                                        'Justificado': '#3498db'
                                    }, hole=0.4)
                st.plotly_chart(fig_pastel, use_container_width=True)
                
            with col_der:
                st.subheader("Comparativa de Asistencia por Cursos")
                asistencia_por_curso = df_final_display.groupby(['Curso', 'Estado']).size().unstack(fill_value=0).reset_index()
                for c in ['P', 'T', 'A', 'J']:
                    if c not in asistencia_por_curso.columns: 
                        asistencia_por_curso[c] = 0
                asistencia_por_curso['Total'] = asistencia_por_curso['P'] + asistencia_por_curso['T'] + asistencia_por_curso['A'] + asistencia_por_curso['J']
                asistencia_por_curso['% Asistencia'] = ((asistencia_por_curso['P'] + asistencia_por_curso['T']) / asistencia_por_curso['Total']) * 100
                
                fig_barra = px.bar(asistencia_por_curso, x='Curso', y='% Asistencia', 
                                   title="Porcentaje de Asistencia por Grado", labels={'% Asistencia': 'Porcentaje (%)'},
                                   color='Curso')
                fig_barra.update_yaxes(range=[60, 101])
                st.plotly_chart(fig_barra, use_container_width=True)
                
            st.markdown("---")
            
            st.subheader("Alerta Temprana: Estudiantes con Mayor Cantidad de Faltas")
            ranking = df_final_display.groupby(['Curso', 'Estudiante', 'Estado_Nombre']).size().unstack(fill_value=0).reset_index()
            if 'Ausente (Falta)' in ranking.columns:
                ranking = ranking.sort_values(by='Ausente (Falta)', ascending=False).head(15)
                ranking = ranking.rename(columns={'Ausente (Falta)': 'Faltas Injustificadas'})
                st.dataframe(ranking[['Curso', 'Estudiante', 'Faltas Injustificadas']], use_container_width=True)
            else:
                st.info("No se registran inasistencias en este grupo.")
        else:
            st.warning("No hay datos para mostrar con los filtros seleccionados.")

except FileNotFoundError:
    st.error(f"No se encontró el archivo '{ARCHIVO_EXCEL}' en la carpeta especificada.")
except Exception as e:
    st.error(f"Ocurrió un error inesperado al procesar los datos: {e}")