import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl

st.set_page_config(page_title="Control de Asistencias 2025-2026", layout="wide")

st.title("Dashboard Asistencias Unidad Educativa Caranqui")
st.markdown("### Periodo Lectivo 2025 - 2026")

# Nombre de tu archivo central
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
    
    # Recorrer cada pestaña del archivo (omitir la de Inicio)
    for hoja in hojas:
        if hoja.lower() == "inicio":
            continue
            
        # Intentar deducir Curso y Año del nombre de la pestaña (Ej: "Octavo A-2025")
        partes = hoja.split('-')
        curso = partes[0].strip()
        año = partes[1].strip() if len(partes) > 1 else "2026"
        
        # Leer la hoja completa
        df_hoja = pd.read_excel(ruta_archivo, sheet_name=hoja, header=None)
        
        mes_actual = None
        
        # Recorrer fila por fila para transformar la matriz en filas individuales
        for idx, row in df_hoja.iterrows():
            valor_celda_1 = str(row.iloc[0]).strip().upper()
            
            # Detectar si la fila marca el inicio de un mes
            if valor_celda_1 in meses_dict:
                mes_actual = valor_celda_1
                continue
            
            # Detectar si es la fila de cabecera de días (N°, Estudiante, 1, 2, 3...)
            if "ESTUDIANTE" in str(row.iloc[1]).upper() or "N°" in valor_celda_1:
                continue
                
            # Si tenemos un mes identificado y la fila tiene un nombre de estudiante válido
            estudiante = row.iloc[1]
            if pd.notna(estudiante) and str(estudiante).strip() != "" and mes_actual:
                if "TOTAL" in str(estudiante).upper() or "ESTUDIANTE" in str(estudiante).upper():
                    continue
                
                # Revisar las columnas de los días 1 al 31 (columnas índice 2 al 32)
                for dia in range(1, 32):
                    col_idx = dia + 1
                    if col_idx < len(row):
                        estado = row.iloc[col_idx]
                        
                        # Si está vacío, asumimos que estuvo 'P' (Presente), de lo contrario registramos el código
                        if pd.isna(estado) or str(estado).strip() == "":
                            estado_final = "P"
                        else:
                            estado_final = str(estado).strip().upper()
                        
                        # Guardar solo códigos válidos para evitar basura del Excel
                        if estado_final in ['P', 'A', 'T', 'J']:
                            lista_registros.append({
                                "Año": año,
                                "Curso": curso,
                                "Estudiante": str(estudiante).strip(),
                                "Mes": mes_actual,
                                "Dia": dia,
                                "Estado": estado_final
                            })
                            
    # Convertir a DataFrame unificado
    df_final = pd.DataFrame(lista_registros)
    
    # Mapear estados a nombres entendibles para los gráficos
    mapa_estados = {'P': 'Presente', 'A': 'Ausente (Falta)', 'T': 'Tarde (Atraso)', 'J': 'Justificado'}
    df_final['Estado_Nombre'] = df_final['Estado'].map(mapa_estados)
    
    return df_final

# Cargar y procesar datos automáticamente
try:
    with st.spinner("Procesando la base de datos inmensa... Esto solo tardará unos segundos."):
        df_asistencias = procesar_todo_el_excel(ARCHIVO_EXCEL)
    
    # --- FILTROS EN BARRA LATERAL ---
    st.sidebar.header("Filtros de Control")
    
    filtro_año = st.sidebar.selectbox("Seleccione el Año Lectivo:", options=df_asistencias['Año'].unique())
    df_filtrado = df_asistencias[df_asistencias['Año'] == filtro_año]
    
    lista_cursos = sorted(df_filtrado['Curso'].unique())
    filtro_curso = st.sidebar.selectbox("Seleccione el Curso / Grado:", options=["Todos los Cursos"] + lista_cursos)
    
    if filtro_curso != "Todos los Cursos":
        df_filtrado = df_filtrado[df_filtrado['Curso'] == filtro_curso]

    # --- CÁLCULO DE MÉTRICAS GENERALES (KPIs) ---
    total_dias_controlados = len(df_filtrado)
    asistencias_efectivas = len(df_filtrado[df_filtrado['Estado'].isin(['P', 'T'])]) # Presente y Tarde cuentan como asistencia
    
    porcentaje_asistencia = (asistencias_efectivas / total_dias_controlados) * 100 if total_dias_controlados > 0 else 0
    total_faltas = len(df_filtrado[df_filtrado['Estado'] == 'A'])
    total_atrasos = len(df_filtrado[df_filtrado['Estado'] == 'T'])

    # Mostrar tarjetas KPIs
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric(label="Tasa de Asistencia Promedio", value=f"{porcentaje_asistencia:.1f}%")
    kpi2.metric(label="Total Inasistencias (Faltas)", value=f"{total_faltas:,}")
    kpi3.metric(label="Total de Atrasos Registrados", value=f"{total_atrasos:,}")

    st.markdown("---")

    # --- SECCIÓN DE GRÁFICOS ---
    col_izq, col_der = st.columns(2)

    with col_izq:
        st.subheader("Distribución General de Estados")
        fig_pastel = px.pie(df_filtrado, names='Estado_Nombre', 
                            color='Estado_Nombre',
                            color_discrete_map={
                                'Presente': '#2ecc71',
                                'Ausente (Falta)': '#e74c3c',
                                'Tarde (Atraso)': '#f1c40f',
                                'Justificado': '#3498db'
                            }, hole=0.4)
        st.plotly_chart(fig_pastel, use_container_width=True)

    with col_der:
        st.subheader("Comportamiento de Asistencia por Mes")
        # Calcular % de asistencia mensual
        asistencia_mensual = df_filtrado.groupby(['Mes', 'Estado']).size().unstack(fill_value=0).reset_index()
        # Asegurar que existan las columnas para evitar errores
        for col in ['P', 'T', 'A', 'J']:
            if col not in asistencia_mensual.columns:
                asistencia_mensual[col] = 0
                
        asistencia_mensual['Total'] = asistencia_mensual['P'] + asistencia_mensual['T'] + asistencia_mensual['A'] + asistencia_mensual['J']
        asistencia_mensual['% Asistencia'] = ((asistencia_mensual['P'] + asistencia_mensual['T']) / asistencia_mensual['Total']) * 100
        
        # Ordenar meses cronológicamente
        orden_meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
        asistencia_mensual['Mes'] = pd.Categorical(asistencia_mensual['Mes'], categories=orden_meses, ordered=True)
        asistencia_mensual = asistencia_mensual.sort_values('Mes')

        fig_linea = px.line(asistencia_mensual, x='Mes', y='% Asistencia', markers=True,
                            title="% de Alumnos Presentes", labels={'% Asistencia': 'Porcentaje (%)'})
        fig_linea.update_yaxes(range=[50, 101])
        st.plotly_chart(fig_linea, use_container_width=True)

    st.markdown("---")

    # --- RANKING / ALERTAS (Crítico para Inspección o Dirección) ---
    st.subheader("Estudiantes con Mayor Nivel de Ausentismo (Alerta Temprana)")
    
    # Calcular ausentismo por estudiante
    ranking_estudiantes = df_filtrado.groupby(['Curso', 'Estudiante', 'Estado']).size().unstack(fill_value=0).reset_index()
    if 'A' in ranking_estudiantes.columns:
        ranking_estudiantes = ranking_estudiantes.sort_values(by='A', ascending=False).head(10)
        ranking_estudiantes = ranking_estudiantes.rename(columns={'P': 'Presentes', 'A': 'Faltas', 'T': 'Atrasos', 'J': 'Justificados'})
        st.dataframe(ranking_estudiantes[['Curso', 'Estudiante', 'Faltas', 'Atrasos']], use_container_width=True)
    else:
        st.info("No se registran faltas en el periodo seleccionado.")

except FileNotFoundError:
    st.error(f"No se encontró el archivo '{ARCHIVO_EXCEL}' en la carpeta actual. Por favor, asegúrate de colocar el archivo Excel con ese nombre exacto junto a este script de Python.")
except Exception as e:
    st.error(f"Ocurrió un error al procesar el archivo: {e}")