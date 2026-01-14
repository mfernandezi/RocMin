import streamlit as st
import pandas as pd
from datetime import datetime, time
from io import BytesIO
from streamlit_drawable_canvas import st_canvas
from PIL import Image
import xlsxwriter
import base64

st.set_page_config(page_title="ROCMIN - Registro de Producción", layout="wide")

st.title("ROCMIN - Registro de Producción Diaria")

# Inicializar estado de sesión
if 'num_pozos' not in st.session_state:
    st.session_state.num_pozos = 3

if 'num_demoras' not in st.session_state:
    st.session_state.num_demoras = {key: 1 for key in range(26)}

# ============================================
# SECCIÓN 1: DATOS DEL TURNO
# ============================================
st.header("Datos del Turno")

col1, col2, col3 = st.columns(3)
with col1:
    operador = st.text_input("Operador *", key="operador")
    equipo = st.text_input("Equipo", key="equipo")
    fecha = st.date_input("Fecha", value=datetime.now(), key="fecha")

with col2:
    jefe_turno = st.text_input("Jefe de Turno", key="jefe_turno")
    tricono = st.text_input("Tricono", key="tricono")
    turno = st.selectbox("Turno", ["Día", "Noche"], key="turno")

with col3:
    st.write("")  # Espaciador

# ============================================
# SECCIÓN 2: HORÓMETROS
# ============================================
st.header("Horómetros")

col1, col2 = st.columns(2)
with col1:
    st.subheader("Motor")
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        motor_entrada = st.number_input("Entrada", min_value=0.0, step=0.1, key="motor_entrada")
    with col_m2:
        motor_salida = st.number_input("Salida", min_value=0.0, step=0.1, key="motor_salida")

with col2:
    st.subheader("Rotación")
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        rotacion_entrada = st.number_input("Entrada", min_value=0.0, step=0.1, key="rotacion_entrada")
    with col_r2:
        rotacion_salida = st.number_input("Salida", min_value=0.0, step=0.1, key="rotacion_salida")

# ============================================
# SECCIÓN 3: PRODUCCIÓN
# ============================================
st.header("Producción")

col_add, col_remove = st.columns([1, 1])
with col_add:
    if st.button("➕ Agregar Pozo"):
        st.session_state.num_pozos += 1
with col_remove:
    if st.button("➖ Quitar Pozo") and st.session_state.num_pozos > 1:
        st.session_state.num_pozos -= 1

# Headers de la tabla de producción
cols_header = st.columns([1, 1, 1, 1.2, 1.2, 1.2, 1.5, 1])
headers = ["Malla", "N° Pozo", "Metros", "Hora Inicio", "Hora Término", "Tiempo Efectivo", "Demora Op. Traslado", "Diámetro"]
for col, header in zip(cols_header, headers):
    col.markdown(f"**{header}**")

# Filas de producción
produccion_data = []
total_metros = 0.0

for i in range(st.session_state.num_pozos):
    cols = st.columns([1, 1, 1, 1.2, 1.2, 1.2, 1.5, 1])

    malla = cols[0].text_input("", key=f"malla_{i}", label_visibility="collapsed")
    num_pozo = cols[1].text_input("", key=f"num_pozo_{i}", label_visibility="collapsed")
    metros = cols[2].number_input("", min_value=0.0, step=0.1, key=f"metros_{i}", label_visibility="collapsed")
    hora_inicio = cols[3].time_input("", value=None, key=f"hora_inicio_{i}", label_visibility="collapsed")
    hora_termino = cols[4].time_input("", value=None, key=f"hora_termino_{i}", label_visibility="collapsed")
    tiempo_efectivo = cols[5].number_input("", min_value=0, step=1, key=f"tiempo_efectivo_{i}", label_visibility="collapsed")
    demora_traslado = cols[6].number_input("", min_value=0, step=1, key=f"demora_traslado_{i}", label_visibility="collapsed")
    diametro = cols[7].text_input("", key=f"diametro_{i}", label_visibility="collapsed")

    total_metros += metros

    produccion_data.append({
        'Malla': malla,
        'N° Pozo': num_pozo,
        'Metros': metros,
        'Hora Inicio': hora_inicio.strftime("%H:%M") if hora_inicio else "",
        'Hora Término': hora_termino.strftime("%H:%M") if hora_termino else "",
        'Tiempo Efectivo': tiempo_efectivo,
        'Demora Op. Traslado': demora_traslado,
        'Diámetro': diametro
    })

st.metric("Total Metros Reales", f"{total_metros:.2f} m")

# ============================================
# SECCIÓN 4: TIEMPOS DE OPERACIÓN
# ============================================
st.header("Tiempos de Operación")

col1, col2 = st.columns(2)
with col1:
    tiempo_efectivo_total = st.number_input("Tiempo Efectivo Total (min)", min_value=0, step=1, key="tiempo_efectivo_total")
with col2:
    tiempo_demoras_op = st.number_input("Tiempo Demoras Operacionales (min)", min_value=0, step=1, key="tiempo_demoras_op")

# ============================================
# SECCIÓN 5: TIEMPOS DE DEMORA CON MOTOR EN MARCHA
# ============================================
st.header("Tiempos de Demora con Motor en Marcha")

# Lista de tipos de demora
tipos_demora = [
    "Mantenimiento Programado",
    "Falla Mec, Elec, Hidr, Neum",
    "Espera de Repuestos",
    "Espera de Mecánicos",
    "Falla GPS/Cámara",
    "Sin Operador",
    "Abastecimiento de Combustible",
    "Abastecimiento de Agua",
    "Espera de Combustible",
    "Espera de Agua",
    "Reservas Operativas",
    "Sin Acceso",
    "Sin Estandarización",
    "Limpieza de Área",
    "Espera de Escolta",
    "Tronadura",
    "Traslado en Cama Baja",
    "Colación",
    "Traslado Colación",
    "Cambio de Turno",
    "Espera Apoyo Cambio Aceros",
    "Condición de Riesgo/Climático",
    "Revisión Equipo",
    "Charla, Reunión",
    "Limpieza de Equipo",
    "Sin Postura"
]

demoras_data = {}

# Crear tabla de demoras con expansor para cada tipo
for idx, tipo in enumerate(tipos_demora):
    with st.expander(f"{tipo}", expanded=False):
        cols_dem = st.columns([1.5, 1.5, 1.5, 1.5, 1])
        cols_dem[0].markdown("**Desde**")
        cols_dem[1].markdown("**Hasta**")
        cols_dem[2].markdown("**Desde 2**")
        cols_dem[3].markdown("**Hasta 2**")
        cols_dem[4].markdown("**Total Min**")

        cols_input = st.columns([1.5, 1.5, 1.5, 1.5, 1])
        desde1 = cols_input[0].time_input("", value=None, key=f"dem_desde1_{idx}", label_visibility="collapsed")
        hasta1 = cols_input[1].time_input("", value=None, key=f"dem_hasta1_{idx}", label_visibility="collapsed")
        desde2 = cols_input[2].time_input("", value=None, key=f"dem_desde2_{idx}", label_visibility="collapsed")
        hasta2 = cols_input[3].time_input("", value=None, key=f"dem_hasta2_{idx}", label_visibility="collapsed")
        total_min = cols_input[4].number_input("", min_value=0, step=1, key=f"dem_total_{idx}", label_visibility="collapsed")

        demoras_data[tipo] = {
            'Desde 1': desde1.strftime("%H:%M") if desde1 else "",
            'Hasta 1': hasta1.strftime("%H:%M") if hasta1 else "",
            'Desde 2': desde2.strftime("%H:%M") if desde2 else "",
            'Hasta 2': hasta2.strftime("%H:%M") if hasta2 else "",
            'Total Min': total_min
        }

# ============================================
# SECCIÓN 6: ABASTECIMIENTOS
# ============================================
st.header("Abastecimientos")

abastecimientos = [
    "1° Carga de Petróleo Turno",
    "2° Carga de Petróleo Turno",
    "Carga de Agua"
]

abastecimientos_data = {}

cols_ab_header = st.columns([2, 1, 1, 1])
cols_ab_header[0].markdown("**Descripción**")
cols_ab_header[1].markdown("**Litros**")
cols_ab_header[2].markdown("**Hora**")
cols_ab_header[3].markdown("**Horómetro**")

for idx, abast in enumerate(abastecimientos):
    cols_ab = st.columns([2, 1, 1, 1])
    cols_ab[0].write(abast)
    litros = cols_ab[1].number_input("", min_value=0.0, step=0.1, key=f"abast_litros_{idx}", label_visibility="collapsed")
    hora_abast = cols_ab[2].time_input("", value=None, key=f"abast_hora_{idx}", label_visibility="collapsed")
    horometro_abast = cols_ab[3].number_input("", min_value=0.0, step=0.1, key=f"abast_horometro_{idx}", label_visibility="collapsed")

    abastecimientos_data[abast] = {
        'Litros': litros,
        'Hora': hora_abast.strftime("%H:%M") if hora_abast else "",
        'Horómetro': horometro_abast
    }

# ============================================
# SECCIÓN 7: DECLARACIÓN DE INCIDENTES
# ============================================
st.header("Declaración de Incidentes")

incidentes_tipo = st.text_area("Incidentes Ambientales/Personas/Equipos", key="incidentes_tipo")
incidentes_desc = st.text_area("Descripción", key="incidentes_desc")
no_conformidad = st.text_area("No Conformidad", key="no_conformidad")

# ============================================
# SECCIÓN 8: OBSERVACIONES
# ============================================
st.header("Observaciones")

observaciones = st.text_area("", key="observaciones", height=100)

# ============================================
# SECCIÓN 9: FIRMAS
# ============================================
st.header("Firmas")

st.markdown("**Instrucciones:** Dibuje su firma en el recuadro correspondiente usando el mouse o el dedo en dispositivos táctiles.")

col_firma1, col_firma2, col_firma3 = st.columns(3)

with col_firma1:
    st.subheader("Firma Operador")
    canvas_operador = st_canvas(
        fill_color="rgba(255, 165, 0, 0.3)",
        stroke_width=2,
        stroke_color="#000000",
        background_color="#ffffff",
        height=150,
        width=300,
        drawing_mode="freedraw",
        key="canvas_operador",
    )

with col_firma2:
    st.subheader("Firma Supervisor ROCMIN")
    canvas_supervisor_rocmin = st_canvas(
        fill_color="rgba(255, 165, 0, 0.3)",
        stroke_width=2,
        stroke_color="#000000",
        background_color="#ffffff",
        height=150,
        width=300,
        drawing_mode="freedraw",
        key="canvas_supervisor_rocmin",
    )

with col_firma3:
    st.subheader("Firma Supervisor Cliente")
    canvas_supervisor_cliente = st_canvas(
        fill_color="rgba(255, 165, 0, 0.3)",
        stroke_width=2,
        stroke_color="#000000",
        background_color="#ffffff",
        height=150,
        width=300,
        drawing_mode="freedraw",
        key="canvas_supervisor_cliente",
    )

# ============================================
# FUNCIÓN PARA GENERAR EXCEL
# ============================================
def generar_excel():
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Registro Producción')

    # Formatos
    titulo_format = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'bg_color': '#4472C4',
        'font_color': 'white',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D9E2F3',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True
    })

    cell_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    label_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'left',
        'valign': 'vcenter',
        'bg_color': '#E2EFDA'
    })

    value_format = workbook.add_format({
        'border': 1,
        'align': 'left',
        'valign': 'vcenter'
    })

    section_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'bg_color': '#FFC000',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    # Ajustar anchos de columna
    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:J', 15)

    row = 0

    # DATOS DEL TURNO
    worksheet.write(row, 0, 'Operador', label_format)
    worksheet.write(row, 1, operador, value_format)
    row += 1
    worksheet.write(row, 0, 'Jefe de Turno', label_format)
    worksheet.write(row, 1, jefe_turno, value_format)
    row += 1
    worksheet.write(row, 0, 'Equipo', label_format)
    worksheet.write(row, 1, equipo, value_format)
    row += 1
    worksheet.write(row, 0, 'Tricono', label_format)
    worksheet.write(row, 1, tricono, value_format)
    row += 1
    worksheet.write(row, 0, 'Fecha', label_format)
    worksheet.write(row, 1, fecha.strftime("%d/%m/%Y"), value_format)
    row += 1
    worksheet.write(row, 0, 'Turno', label_format)
    worksheet.write(row, 1, turno, value_format)
    row += 2

    # HORÓMETROS
    worksheet.merge_range(row, 0, row, 2, 'Horómetros', section_format)
    row += 1
    worksheet.write(row, 0, '', header_format)
    worksheet.write(row, 1, 'Entrada', header_format)
    worksheet.write(row, 2, 'Salida', header_format)
    row += 1
    worksheet.write(row, 0, 'Motor', label_format)
    worksheet.write(row, 1, motor_entrada, cell_format)
    worksheet.write(row, 2, motor_salida, cell_format)
    row += 1
    worksheet.write(row, 0, 'Rotación', label_format)
    worksheet.write(row, 1, rotacion_entrada, cell_format)
    worksheet.write(row, 2, rotacion_salida, cell_format)
    row += 2

    # PRODUCCIÓN
    worksheet.merge_range(row, 0, row, 7, '*Producción*', section_format)
    row += 1

    headers_prod = ['Malla', 'N° Pozo', 'Metros', 'Hora Inicio', 'Hora Término', 'Tiempo Efectivo', 'Demora Op.', 'Diámetro']
    for col, header in enumerate(headers_prod):
        worksheet.write(row, col, header, header_format)
    row += 1

    for pozo in produccion_data:
        worksheet.write(row, 0, pozo['Malla'], cell_format)
        worksheet.write(row, 1, pozo['N° Pozo'], cell_format)
        worksheet.write(row, 2, pozo['Metros'], cell_format)
        worksheet.write(row, 3, pozo['Hora Inicio'], cell_format)
        worksheet.write(row, 4, pozo['Hora Término'], cell_format)
        worksheet.write(row, 5, pozo['Tiempo Efectivo'], cell_format)
        worksheet.write(row, 6, pozo['Demora Op. Traslado'], cell_format)
        worksheet.write(row, 7, pozo['Diámetro'], cell_format)
        row += 1

    # Total metros
    worksheet.write(row, 0, 'TOTAL METROS REALES', label_format)
    worksheet.write(row, 2, total_metros, cell_format)
    row += 2

    # TIEMPOS DE OPERACIÓN
    worksheet.merge_range(row, 0, row, 1, '*Tiempos de Operación*', section_format)
    row += 1
    worksheet.write(row, 0, 'Tiempo Efectivo Total', label_format)
    worksheet.write(row, 1, tiempo_efectivo_total, cell_format)
    row += 1
    worksheet.write(row, 0, 'Tiempo Demoras Operacionales', label_format)
    worksheet.write(row, 1, tiempo_demoras_op, cell_format)
    row += 2

    # TIEMPOS DE DEMORA CON MOTOR EN MARCHA
    worksheet.merge_range(row, 0, row, 5, '*Tiempos de Demora con Motor en Marcha*', section_format)
    row += 1

    worksheet.write(row, 0, 'Item', header_format)
    worksheet.write(row, 1, 'Desde', header_format)
    worksheet.write(row, 2, 'Hasta', header_format)
    worksheet.write(row, 3, 'Desde', header_format)
    worksheet.write(row, 4, 'Hasta', header_format)
    worksheet.write(row, 5, 'Total Min', header_format)
    row += 1

    for tipo, datos in demoras_data.items():
        worksheet.write(row, 0, tipo, label_format)
        worksheet.write(row, 1, datos['Desde 1'], cell_format)
        worksheet.write(row, 2, datos['Hasta 1'], cell_format)
        worksheet.write(row, 3, datos['Desde 2'], cell_format)
        worksheet.write(row, 4, datos['Hasta 2'], cell_format)
        worksheet.write(row, 5, datos['Total Min'], cell_format)
        row += 1

    row += 1

    # ABASTECIMIENTOS
    worksheet.merge_range(row, 0, row, 3, '*Abastecimientos*', section_format)
    row += 1

    worksheet.write(row, 0, 'Descripción', header_format)
    worksheet.write(row, 1, 'Litros', header_format)
    worksheet.write(row, 2, 'Hora', header_format)
    worksheet.write(row, 3, 'Horómetro', header_format)
    row += 1

    for abast, datos in abastecimientos_data.items():
        worksheet.write(row, 0, abast, label_format)
        worksheet.write(row, 1, datos['Litros'], cell_format)
        worksheet.write(row, 2, datos['Hora'], cell_format)
        worksheet.write(row, 3, datos['Horómetro'], cell_format)
        row += 1

    row += 1

    # DECLARACIÓN DE INCIDENTES
    worksheet.merge_range(row, 0, row, 3, '*Declaración de Incidentes*', section_format)
    row += 1
    worksheet.write(row, 0, 'Incidentes Ambientales/Personas/Equipos', label_format)
    worksheet.merge_range(row, 1, row, 3, incidentes_tipo, value_format)
    row += 1
    worksheet.write(row, 0, 'Descripción', label_format)
    worksheet.merge_range(row, 1, row, 3, incidentes_desc, value_format)
    row += 1
    worksheet.write(row, 0, 'No Conformidad', label_format)
    worksheet.merge_range(row, 1, row, 3, no_conformidad, value_format)
    row += 2

    # OBSERVACIONES
    worksheet.merge_range(row, 0, row, 3, '*Observaciones*', section_format)
    row += 1
    worksheet.merge_range(row, 0, row + 2, 3, observaciones, value_format)
    row += 4

    # FIRMAS
    worksheet.merge_range(row, 0, row, 5, '*Firmas*', section_format)
    row += 1

    firma_row = row

    # Insertar firmas como imágenes
    if canvas_operador.image_data is not None:
        img = Image.fromarray(canvas_operador.image_data.astype('uint8'), 'RGBA')
        img_byte_arr = BytesIO()
        img.save(img_byte_arr, format='PNG')
        img_byte_arr.seek(0)
        worksheet.write(firma_row, 0, 'Firma Operador:', label_format)
        worksheet.insert_image(firma_row + 1, 0, 'firma_operador.png',
                             {'image_data': img_byte_arr, 'x_scale': 0.5, 'y_scale': 0.5})

    if canvas_supervisor_rocmin.image_data is not None:
        img = Image.fromarray(canvas_supervisor_rocmin.image_data.astype('uint8'), 'RGBA')
        img_byte_arr = BytesIO()
        img.save(img_byte_arr, format='PNG')
        img_byte_arr.seek(0)
        worksheet.write(firma_row, 2, 'Firma Supervisor ROCMIN:', label_format)
        worksheet.insert_image(firma_row + 1, 2, 'firma_supervisor_rocmin.png',
                             {'image_data': img_byte_arr, 'x_scale': 0.5, 'y_scale': 0.5})

    if canvas_supervisor_cliente.image_data is not None:
        img = Image.fromarray(canvas_supervisor_cliente.image_data.astype('uint8'), 'RGBA')
        img_byte_arr = BytesIO()
        img.save(img_byte_arr, format='PNG')
        img_byte_arr.seek(0)
        worksheet.write(firma_row, 4, 'Firma Supervisor Cliente:', label_format)
        worksheet.insert_image(firma_row + 1, 4, 'firma_supervisor_cliente.png',
                             {'image_data': img_byte_arr, 'x_scale': 0.5, 'y_scale': 0.5})

    workbook.close()
    output.seek(0)
    return output

# ============================================
# BOTÓN DE DESCARGA
# ============================================
st.markdown("---")
st.header("Descargar Registro")

if st.button("Generar Excel para Descarga", type="primary"):
    if not operador:
        st.error("Por favor, ingrese el nombre del Operador.")
    else:
        # Generar nombre del archivo: fecha_Turno_NombreOperador
        fecha_str = fecha.strftime("%Y%m%d")
        turno_str = "Dia" if turno == "Día" else "Noche"
        operador_str = operador.replace(" ", "_")
        nombre_archivo = f"{fecha_str}_{turno_str}_{operador_str}.xlsx"

        excel_file = generar_excel()

        st.download_button(
            label=f"📥 Descargar: {nombre_archivo}",
            data=excel_file,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success(f"Excel generado correctamente. Total Metros Reales: {total_metros:.2f} m")

# ============================================
# FOOTER
# ============================================
st.markdown("---")
st.markdown("*ROCMIN - Sistema de Registro de Producción Diaria*")
