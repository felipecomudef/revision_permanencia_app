import streamlit as st
import pandas as pd
import json
from datetime import datetime
import os
from pathlib import Path

# Configuración de la página
st.set_page_config(page_title="Control de Asistencia", layout="wide")
st.title("📋 Control de Asistencia por Centro de Costo")

# Causales de ausencia predefinidas
CAUSALES_AUSENCIA = [
    "Feriado",
    "Licencia médica",
    "Licencia médica prolongada",
    "Licencia médica reciente",
    "Vacaciones",
    "Permiso administrativo",
    "Suspensión",
    "Suspensión de funciones por procedimiento administrativo",
    "No se presentó",
    "Otra razón"
]

# Crear directorios si no existen
if not os.path.exists("datos_asistencia"):
    os.makedirs("datos_asistencia")
if not os.path.exists("archivo_maestro"):
    os.makedirs("archivo_maestro")

# Función para guardar datos
def guardar_datos(datos, fecha, responsable):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"datos_asistencia/revision_{fecha}_{responsable}_{timestamp}.json"
    with open(nombre_archivo, 'w', encoding='utf-8') as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)
    return nombre_archivo

# Función para convertir datos a Excel para descarga
def datos_a_excel(datos, df_original, fecha, responsable):
    resultados = []
    for item in datos:
        resultados.append({
            'Fecha': fecha,
            'Responsable': responsable,
            'RUN': item['RUN'],
            'Apellido Paterno': item['Apellido Paterno'],
            'Apellido Materno': item['Apellido Materno'],
            'Nombres': item['Nombres'],
            'Centro Costo': item['Centro Costo'],
            'Cargo': item['Cargo'],
            'Presente': 'Sí' if item['presente'] else 'No',
            'Causal Ausencia': item.get('causal_ausencia', ''),
            'Comentarios': item.get('comentarios', '')
        })
    return pd.DataFrame(resultados)

# Función para obtener archivo maestro guardado
def obtener_archivo_maestro():
    if not os.path.exists("archivo_maestro"):
        os.makedirs("archivo_maestro")
    archivos = list(Path("archivo_maestro").glob("*.xlsx")) + list(Path("archivo_maestro").glob("*.xls"))
    if archivos:
        return archivos[0]
    return None

# Función para listar revisiones guardadas
def listar_revisiones():
    archivos = list(Path("datos_asistencia").glob("*.json"))
    revisiones = []
    for archivo in archivos:
        try:
            with open(archivo, 'r', encoding='utf-8') as f:
                datos = json.load(f)
                if datos:
                    revisiones.append({
                        'archivo': archivo.name,
                        'ruta': str(archivo),
                        'fecha': datos[0].get('Centro Costo', 'N/A'),
                        'centro': datos[0].get('Centro Costo', 'N/A')
                    })
        except:
            pass
    return sorted(revisiones, key=lambda x: x['archivo'], reverse=True)

# Pestañas principales
tab1, tab2, tab3 = st.tabs(["📝 Nueva Revisión", "📊 Historial", "⚙️ Configuración"])

# ============= TAB 1: NUEVA REVISIÓN =============
with tab1:
    st.header("Registrar Nueva Revisión")
    
    # Obtener archivo maestro
    archivo_maestro = obtener_archivo_maestro()
    
    if archivo_maestro:
        df = pd.read_excel(archivo_maestro)
        
        # Validar columnas requeridas
        columnas_requeridas = ['RUN', 'Apellido Paterno', 'Apellido Materno', 
                              'Nombres', 'Centro Costo', 'Cargo']
        if all(col in df.columns for col in columnas_requeridas):
            
            # Sección de información de la revisión
            col1, col2 = st.columns(2)
            
            with col1:
                fecha_revision = st.date_input("Fecha de la revisión", value=datetime.now())
            
            with col2:
                responsable = st.text_input("Responsable de la revisión")
            
            # Filtro por centro de costo
            centros = sorted(df['Centro Costo'].unique().tolist())
            centro_seleccionado = st.selectbox("Filtrar por Centro de Costo", centros)
            
            # Filtrar datos
            df_filtrado = df[df['Centro Costo'] == centro_seleccionado].reset_index(drop=True)
            
            st.subheader(f"Centro de Costo: {centro_seleccionado}")
            st.write(f"Total de funcionarios: {len(df_filtrado)}")
            
            # Inicializar estado de sesión para guardar datos
            if 'datos_revision' not in st.session_state or st.session_state.get('centro_anterior') != centro_seleccionado:
                st.session_state.datos_revision = []
                st.session_state.centro_anterior = centro_seleccionado
                for idx, row in df_filtrado.iterrows():
                    st.session_state.datos_revision.append({
                        'RUN': row['RUN'],
                        'Apellido Paterno': row['Apellido Paterno'],
                        'Apellido Materno': row['Apellido Materno'],
                        'Nombres': row['Nombres'],
                        'Centro Costo': row['Centro Costo'],
                        'Cargo': row['Cargo'],
                        'presente': False,
                        'causal_ausencia': '',
                        'comentarios': ''
                    })
            
            # Tabla interactiva
            st.subheader("Registro de Asistencia")
            
            for idx in range(len(st.session_state.datos_revision)):
                item = st.session_state.datos_revision[idx]
                
                col1, col2, col3, col4, col5 = st.columns([2, 1, 2, 2, 3])
                
                with col1:
                    st.write(f"**{item['Nombres']} {item['Apellido Paterno']}**")
                    st.caption(f"RUN: {item['RUN']}")
                
                with col2:
                    presente = st.checkbox(
                        "Presente",
                        value=item['presente'],
                        key=f"presente_{idx}"
                    )
                    st.session_state.datos_revision[idx]['presente'] = presente
                
                with col3:
                    if not presente:
                        causal = st.selectbox(
                            "Causal",
                            CAUSALES_AUSENCIA,
                            index=0 if not item['causal_ausencia'] else 
                                  CAUSALES_AUSENCIA.index(item['causal_ausencia']) 
                                  if item['causal_ausencia'] in CAUSALES_AUSENCIA else 0,
                            key=f"causal_{idx}",
                            label_visibility="collapsed"
                        )
                        st.session_state.datos_revision[idx]['causal_ausencia'] = causal
                    else:
                        st.session_state.datos_revision[idx]['causal_ausencia'] = ''
                
                with col4:
                    st.write(f"*{item['Cargo']}*")
                
                with col5:
                    comentario = st.text_input(
                        "Comentario",
                        value=item['comentarios'],
                        key=f"comentario_{idx}",
                        label_visibility="collapsed",
                        placeholder="Agregar comentario..."
                    )
                    st.session_state.datos_revision[idx]['comentarios'] = comentario
                
                st.divider()
            
            # Botones de acción
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if st.button("💾 Guardar Revisión", use_container_width=True):
                    if responsable:
                        archivo_guardado = guardar_datos(
                            st.session_state.datos_revision,
                            fecha_revision.strftime("%Y%m%d"),
                            responsable
                        )
                        st.success(f"✅ Datos guardados correctamente")
                    else:
                        st.error("⚠️ Ingresa el nombre del responsable")
            
            with col2:
                if responsable:
                    df_descarga = datos_a_excel(st.session_state.datos_revision, df_filtrado, fecha_revision.strftime("%Y-%m-%d"), responsable)
                    csv = df_descarga.to_csv(index=False, encoding='utf-8-sig')
                    st.download_button(
                        label="📥 Descargar CSV",
                        data=csv,
                        file_name=f"asistencia_{centro_seleccionado}_{fecha_revision}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                else:
                    st.info("Ingresa responsable para descargar")
            
            with col3:
                if responsable:
                    df_descarga = datos_a_excel(st.session_state.datos_revision, df_filtrado, fecha_revision.strftime("%Y-%m-%d"), responsable)
                    excel_buffer = pd.ExcelWriter('temp.xlsx', engine='openpyxl')
                    df_descarga.to_excel(excel_buffer, index=False, sheet_name='Asistencia')
                    excel_buffer.close()
                    
                    with open('temp.xlsx', 'rb') as f:
                        st.download_button(
                            label="📊 Descargar Excel",
                            data=f.read(),
                            file_name=f"asistencia_{centro_seleccionado}_{fecha_revision}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                else:
                    st.info("Ingresa responsable para descargar")
            
            # Estadísticas
            st.subheader("Estadísticas")
            col1, col2, col3 = st.columns(3)
            
            datos_centro_actual = [item for item in st.session_state.datos_revision if item['Centro Costo'] == centro_seleccionado]
            presentes = sum(1 for item in datos_centro_actual if item['presente'])
            ausentes = len(datos_centro_actual) - presentes
            
            with col1:
                st.metric("Presentes", presentes)
            with col2:
                st.metric("Ausentes", ausentes)
            with col3:
                porcentaje = (presentes / len(datos_centro_actual) * 100) if datos_centro_actual else 0
                st.metric("% Asistencia", f"{porcentaje:.1f}%")
        
        else:
            st.error("⚠️ El archivo no contiene todas las columnas requeridas")
            st.info(f"Columnas requeridas: {', '.join(columnas_requeridas)}")
    
    else:
        st.warning("⚠️ No hay archivo maestro cargado. Ve a la pestaña 'Configuración' para cargar uno.")

# ============= TAB 2: HISTORIAL =============
with tab2:
    st.header("Historial de Revisiones")
    
    revisiones = listar_revisiones()
    
    if revisiones:
        # Filtro por centro de costo
        centros_unicos = set()
        for rev in revisiones:
            try:
                with open(rev['ruta'], 'r', encoding='utf-8') as f:
                    datos = json.load(f)
                    if datos:
                        centros_unicos.add(datos[0]['Centro Costo'])
            except:
                pass
        
        centros_filtro = sorted(list(centros_unicos))
        if centros_filtro:
            centro_filtro = st.selectbox("Filtrar por Centro de Costo", ["Todos"] + centros_filtro)
        else:
            centro_filtro = "Todos"
        
        st.subheader(f"Total de revisiones: {len(revisiones)}")
        st.divider()
        
        # Mostrar revisiones
        for rev in revisiones:
            try:
                with open(rev['ruta'], 'r', encoding='utf-8') as f:
                    datos = json.load(f)
                    
                    if not datos:
                        continue
                    
                    centro = datos[0]['Centro Costo']
                    responsable = rev['archivo'].split('_')[2] if len(rev['archivo'].split('_')) > 2 else 'N/A'
                    fecha = rev['archivo'].split('_')[1]
                    
                    # Filtrar por centro
                    if centro_filtro != "Todos" and centro != centro_filtro:
                        continue
                    
                    # Información de la revisión
                    col1, col2, col3, col4 = st.columns([2, 2, 2, 1])
                    
                    with col1:
                        st.write(f"**Centro:** {centro}")
                    with col2:
                        st.write(f"**Responsable:** {responsable}")
                    with col3:
                        st.write(f"**Fecha:** {fecha[:4]}-{fecha[4:6]}-{fecha[6:8]}")
                    
                    # Botón de descarga
                    with col4:
                        # Convertir datos a DataFrame
                        df_historial = pd.DataFrame(datos)
                        csv = df_historial.to_csv(index=False, encoding='utf-8-sig')
                        st.download_button(
                            label="📥",
                            data=csv,
                            file_name=f"{rev['archivo'].replace('.json', '.csv')}",
                            mime="text/csv",
                            key=f"download_{rev['archivo']}"
                        )
                    
                    # Estadísticas rápidas
                    presentes_hist = sum(1 for item in datos if item.get('presente', False))
                    ausentes_hist = len(datos) - presentes_hist
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Presentes", presentes_hist, label_visibility="collapsed")
                    with col2:
                        st.metric("Ausentes", ausentes_hist, label_visibility="collapsed")
                    with col3:
                        porcentaje_hist = (presentes_hist / len(datos) * 100) if datos else 0
                        st.metric("% Asistencia", f"{porcentaje_hist:.1f}%", label_visibility="collapsed")
                    
                    st.divider()
            
            except Exception as e:
                st.error(f"Error al leer {rev['archivo']}: {str(e)}")
    
    else:
        st.info("📭 No hay revisiones guardadas aún.")

# ============= TAB 3: CONFIGURACIÓN =============
with tab3:
    st.header("Configuración")
    
    st.subheader("📂 Archivo Maestro (Listado de Funcionarios)")
    
    archivo_actual = obtener_archivo_maestro()
    if archivo_actual:
        st.success(f"✅ Archivo cargado: {archivo_actual.name}")
        
        if st.button("🗑️ Eliminar archivo actual", use_container_width=True):
            os.remove(archivo_actual)
            st.success("Archivo eliminado. Recarga la página.")
            st.rerun()
    
    uploaded_file = st.file_uploader("Cargar nuevo archivo Excel", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Guardar archivo
        ruta_guardada = f"archivo_maestro/{uploaded_file.name}"
        with open(ruta_guardada, 'wb') as f:
            f.write(uploaded_file.getbuffer())
        
        st.success(f"✅ Archivo cargado: {uploaded_file.name}")
        st.info("Ahora puedes ir a la pestaña 'Nueva Revisión' para comenzar")
        
        # Preview del archivo
        df_preview = pd.read_excel(ruta_guardada)
        st.subheader("Preview del archivo:")
        st.dataframe(df_preview.head(10))
    
    st.divider()
    
    st.subheader("🗂️ Gestionar Revisiones Guardadas")
    
    revisiones_totales = listar_revisiones()
    st.write(f"Total de revisiones guardadas: {len(revisiones_totales)}")
    
    if revisiones_totales and st.button("🗑️ Eliminar todas las revisiones", use_container_width=True):
        for rev in revisiones_totales:
            try:
                os.remove(rev['ruta'])
            except:
                pass
        st.success("Todas las revisiones han sido eliminadas")
        st.rerun()