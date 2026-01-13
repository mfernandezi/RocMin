import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
from io import BytesIO
from PIL import Image
import xlsxwriter

st.set_page_config(page_title="ROCMIN - Registro de Produccion", layout="wide")

st.title("ROCMIN - Registro de Produccion Diaria")

# Inicializar estado de sesion
if 'num_pozos' not in st.session_state:
    st.session_state.num_pozos = 3

# Lista de tipos de demora
tipos_demora = [
    "Mantenimiento Programado",
    "Falla Mec, Elec, Hidr, Neum",
    "Espera de Repuestos",
    "Espera de Mecanicos",
    "Falla GPS/Camara",
    "Sin Operador",
    "Abastecimiento de Combustible",
    "Abastecimiento de Agua",
    "Espera de Combustible",
    "Espera de Agua",
    "Reservas Operativas",
    "Sin Acceso",
    "Sin Estandarizacion",
    "Limpieza de Area",
    "Espera de Escolta",
    "Tronadura",
    "Traslado en Cama Baja",
    "Colacion",
    "Traslado Colacion",
    "Cambio de Turno",
    "Espera Apoyo Cambio Aceros",
    "Condicion de Riesgo/Climatico",
    "Revision Equipo",
    "Charla, Reunion",
    "Limpieza de Equipo",
    "Sin Postura"
]

# Inicializar contadores de filas para cada tipo de demora
for idx in range(len(tipos_demora)):
    if f'num_filas_demora_{idx}' not in st.session_state:
        st.session_state[f'num_filas_demora_{idx}'] = 1

# ============================================
# SECCION 1: DATOS DEL TURNO
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
    turno = st.selectbox("Turno", ["Dia", "Noche"], key="turno")

with col3:
    st.write("")  # Espaciador

# ============================================
# SECCION 2: HOROMETROS
# ============================================
st.header("Horometros")

col1, col2 = st.columns(2)
with col1:
    st.subheader("Motor")
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        motor_entrada = st.number_input("Entrada", min_value=0.0, step=0.1, key="motor_entrada")
    with col_m2:
        motor_salida = st.number_input("Salida", min_value=0.0, step=0.1, key="motor_salida")

with col2:
    st.subheader("Rotacion")
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        rotacion_entrada = st.number_input("Entrada", min_value=0.0, step=0.1, key="rotacion_entrada")
    with col_r2:
        rotacion_salida = st.number_input("Salida", min_value=0.0, step=0.1, key="rotacion_salida")

# ============================================
# SECCION 3: PRODUCCION
# ============================================
st.header("Produccion")

col_add, col_remove = st.columns([1, 1])
with col_add:
    if st.button("+ Agregar Pozo"):
        st.session_state.num_pozos += 1
with col_remove:
    if st.button("- Quitar Pozo") and st.session_state.num_pozos > 1:
        st.session_state.num_pozos -= 1

# Headers de la tabla de produccion
cols_header = st.columns([1, 1, 1, 1.2, 1.2, 1.2, 1.5, 1])
headers = ["Malla", "N Pozo", "Metros", "Hora Inicio", "Hora Termino", "Tiempo Efectivo", "Demora Op. Traslado", "Diametro"]
for col, header in zip(cols_header, headers):
    col.markdown(f"**{header}**")

# Filas de produccion
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
        'N Pozo': num_pozo,
        'Metros': metros,
        'Hora Inicio': hora_inicio.strftime("%H:%M") if hora_inicio else "",
        'Hora Termino': hora_termino.strftime("%H:%M") if hora_termino else "",
        'Tiempo Efectivo': tiempo_efectivo,
        'Demora Op. Traslado': demora_traslado,
        'Diametro': diametro
    })

st.metric("Total Metros Reales", f"{total_metros:.2f} m")

# ============================================
# SECCION 4: TIEMPOS DE OPERACION
# ============================================
st.header("Tiempos de Operacion")

col1, col2 = st.columns(2)
with col1:
    tiempo_efectivo_total = st.number_input("Tiempo Efectivo Total (min)", min_value=0, step=1, key="tiempo_efectivo_total")
with col2:
    tiempo_demoras_op = st.number_input("Tiempo Demoras Operacionales (min)", min_value=0, step=1, key="tiempo_demoras_op")

# ============================================
# SECCION 5: TIEMPOS DE DEMORA CON MOTOR EN MARCHA
# ============================================
st.header("Tiempos de Demora con Motor en Marcha")

def parse_hora(hora_str):
    """Convierte string HH:MM a objeto time, retorna None si es invalido"""
    if not hora_str or hora_str.strip() == "":
        return None
    try:
        # Intentar parsear formato HH:MM
        hora_str = hora_str.strip()
        if ":" in hora_str:
            partes = hora_str.split(":")
            h = int(partes[0])
            m = int(partes[1]) if len(partes) > 1 else 0
            if 0 <= h <= 23 and 0 <= m <= 59:
                return time(h, m)
    except:
        pass
    return None

def calcular_minutos(hora_desde_str, hora_hasta_str):
    """Calcula la diferencia en minutos entre dos horas (strings HH:MM)"""
    hora_desde = parse_hora(hora_desde_str)
    hora_hasta = parse_hora(hora_hasta_str)

    if hora_desde is None or hora_hasta is None:
        return 0

    # Convertir a datetime para poder restar
    hoy = datetime.today().date()
    dt_desde = datetime.combine(hoy, hora_desde)
    dt_hasta = datetime.combine(hoy, hora_hasta)

    # Si hasta es menor que desde, asumimos que cruza medianoche
    if dt_hasta < dt_desde:
        dt_hasta += timedelta(days=1)

    diff = dt_hasta - dt_desde
    return int(diff.total_seconds() / 60)

demoras_data = {}

# Crear tabla de demoras con expansor para cada tipo
for idx, tipo in enumerate(tipos_demora):
    with st.expander(f"{tipo}", expanded=False):
        # Botones para agregar/quitar filas
        col_btn1, col_btn2, col_info = st.columns([1, 1, 2])
        with col_btn1:
            if st.button(f"+ Agregar fila", key=f"add_fila_{idx}"):
                st.session_state[f'num_filas_demora_{idx}'] += 1
                st.rerun()
        with col_btn2:
            if st.button(f"- Quitar fila", key=f"remove_fila_{idx}") and st.session_state[f'num_filas_demora_{idx}'] > 1:
                st.session_state[f'num_filas_demora_{idx}'] -= 1
                st.rerun()

        # Headers
        cols_header = st.columns([2, 2, 1.5])
        cols_header[0].markdown("**Desde (HH:MM)**")
        cols_header[1].markdown("**Hasta (HH:MM)**")
        cols_header[2].markdown("**Minutos**")

        total_minutos_tipo = 0
        filas_demora = []

        # Filas dinamicas
        for fila in range(st.session_state[f'num_filas_demora_{idx}']):
            cols_input = st.columns([2, 2, 1.5])
            desde = cols_input[0].text_input("", key=f"dem_desde_{idx}_{fila}", placeholder="ej: 08:30", label_visibility="collapsed")
            hasta = cols_input[1].text_input("", key=f"dem_hasta_{idx}_{fila}", placeholder="ej: 09:15", label_visibility="collapsed")

            # Calcular minutos automaticamente
            minutos = calcular_minutos(desde, hasta)
            total_minutos_tipo += minutos

            # Mostrar minutos calculados
            if desde and hasta and minutos == 0:
                cols_input[2].markdown(":red[Formato invalido]")
            else:
                cols_input[2].markdown(f"**{minutos}** min")

            filas_demora.append({
                'Desde': desde if desde else "",
                'Hasta': hasta if hasta else "",
                'Minutos': minutos
            })

        # Mostrar total
        st.markdown(f"### Total: **{total_minutos_tipo}** minutos")

        demoras_data[tipo] = {
            'filas': filas_demora,
            'total_min': total_minutos_tipo
        }

# ============================================
# SECCION 6: ABASTECIMIENTOS
# ============================================
st.header("Abastecimientos")

abastecimientos = [
    "1 Carga de Petroleo Turno",
    "2 Carga de Petroleo Turno",
    "Carga de Agua"
]

abastecimientos_data = {}

cols_ab_header = st.columns([2, 1, 1, 1])
cols_ab_header[0].markdown("**Descripcion**")
cols_ab_header[1].markdown("**Litros**")
cols_ab_header[2].markdown("**Hora**")
cols_ab_header[3].markdown("**Horometro**")

for idx, abast in enumerate(abastecimientos):
    cols_ab = st.columns([2, 1, 1, 1])
    cols_ab[0].write(abast)
    litros = cols_ab[1].number_input("", min_value=0.0, step=0.1, key=f"abast_litros_{idx}", label_visibility="collapsed")
    hora_abast = cols_ab[2].time_input("", value=None, key=f"abast_hora_{idx}", label_visibility="collapsed")
    horometro_abast = cols_ab[3].number_input("", min_value=0.0, step=0.1, key=f"abast_horometro_{idx}", label_visibility="collapsed")

    abastecimientos_data[abast] = {
        'Litros': litros,
        'Hora': hora_abast.strftime("%H:%M") if hora_abast else "",
        'Horometro': horometro_abast
    }

# ============================================
# SECCION 7: DECLARACION DE INCIDENTES
# ============================================
st.header("Declaracion de Incidentes")

incidentes_tipo = st.text_area("Incidentes Ambientales/Personas/Equipos", key="incidentes_tipo")
incidentes_desc = st.text_area("Descripcion", key="incidentes_desc")
no_conformidad = st.text_area("No Conformidad", key="no_conformidad")

# ============================================
# SECCION 8: OBSERVACIONES
# ============================================
st.header("Observaciones")

observaciones = st.text_area("", key="observaciones", height=100)

# ============================================
# SECCION 9: FIRMAS (Subida de imagenes)
# ============================================
st.header("Firmas")

st.markdown("**Instrucciones:** Suba una imagen de su firma (PNG, JPG) o tome una foto de su firma.")

col_firma1, col_firma2, col_firma3 = st.columns(3)

with col_firma1:
    st.subheader("Firma Operador")
    firma_operador_file = st.file_uploader("Subir firma", type=['png', 'jpg', 'jpeg'], key="firma_operador")
    if firma_operador_file:
        st.image(firma_operador_file, width=200)

with col_firma2:
    st.subheader("Firma Supervisor ROCMIN")
    firma_supervisor_rocmin_file = st.file_uploader("Subir firma", type=['png', 'jpg', 'jpeg'], key="firma_supervisor_rocmin")
    if firma_supervisor_rocmin_file:
        st.image(firma_supervisor_rocmin_file, width=200)

with col_firma3:
    st.subheader("Firma Supervisor Cliente")
    firma_supervisor_cliente_file = st.file_uploader("Subir firma", type=['png', 'jpg', 'jpeg'], key="firma_supervisor_cliente")
    if firma_supervisor_cliente_file:
        st.image(firma_supervisor_cliente_file, width=200)

# ============================================
# FUNCION PARA GENERAR EXCEL
# ============================================
def generar_excel():
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Registro Produccion')

    # Formatos
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

    total_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#FFEB9C'
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

    # HOROMETROS
    worksheet.merge_range(row, 0, row, 2, 'Horometros', section_format)
    row += 1
    worksheet.write(row, 0, '', header_format)
    worksheet.write(row, 1, 'Entrada', header_format)
    worksheet.write(row, 2, 'Salida', header_format)
    row += 1
    worksheet.write(row, 0, 'Motor', label_format)
    worksheet.write(row, 1, motor_entrada, cell_format)
    worksheet.write(row, 2, motor_salida, cell_format)
    row += 1
    worksheet.write(row, 0, 'Rotacion', label_format)
    worksheet.write(row, 1, rotacion_entrada, cell_format)
    worksheet.write(row, 2, rotacion_salida, cell_format)
    row += 2

    # PRODUCCION
    worksheet.merge_range(row, 0, row, 7, '*Produccion*', section_format)
    row += 1

    headers_prod = ['Malla', 'N Pozo', 'Metros', 'Hora Inicio', 'Hora Termino', 'Tiempo Efectivo', 'Demora Op.', 'Diametro']
    for col, header in enumerate(headers_prod):
        worksheet.write(row, col, header, header_format)
    row += 1

    for pozo in produccion_data:
        worksheet.write(row, 0, pozo['Malla'], cell_format)
        worksheet.write(row, 1, pozo['N Pozo'], cell_format)
        worksheet.write(row, 2, pozo['Metros'], cell_format)
        worksheet.write(row, 3, pozo['Hora Inicio'], cell_format)
        worksheet.write(row, 4, pozo['Hora Termino'], cell_format)
        worksheet.write(row, 5, pozo['Tiempo Efectivo'], cell_format)
        worksheet.write(row, 6, pozo['Demora Op. Traslado'], cell_format)
        worksheet.write(row, 7, pozo['Diametro'], cell_format)
        row += 1

    # Total metros
    worksheet.write(row, 0, 'TOTAL METROS REALES', label_format)
    worksheet.write(row, 2, total_metros, total_format)
    row += 2

    # TIEMPOS DE OPERACION
    worksheet.merge_range(row, 0, row, 1, '*Tiempos de Operacion*', section_format)
    row += 1
    worksheet.write(row, 0, 'Tiempo Efectivo Total', label_format)
    worksheet.write(row, 1, tiempo_efectivo_total, cell_format)
    row += 1
    worksheet.write(row, 0, 'Tiempo Demoras Operacionales', label_format)
    worksheet.write(row, 1, tiempo_demoras_op, cell_format)
    row += 2

    # TIEMPOS DE DEMORA CON MOTOR EN MARCHA
    worksheet.merge_range(row, 0, row, 3, '*Tiempos de Demora con Motor en Marcha*', section_format)
    row += 1

    worksheet.write(row, 0, 'Tipo de Demora', header_format)
    worksheet.write(row, 1, 'Desde', header_format)
    worksheet.write(row, 2, 'Hasta', header_format)
    worksheet.write(row, 3, 'Minutos', header_format)
    row += 1

    for tipo, datos in demoras_data.items():
        if datos['total_min'] > 0:  # Solo mostrar si tiene datos
            first_row = True
            for fila in datos['filas']:
                if fila['Desde'] or fila['Hasta']:
                    if first_row:
                        worksheet.write(row, 0, tipo, label_format)
                        first_row = False
                    else:
                        worksheet.write(row, 0, '', cell_format)
                    worksheet.write(row, 1, fila['Desde'], cell_format)
                    worksheet.write(row, 2, fila['Hasta'], cell_format)
                    worksheet.write(row, 3, fila['Minutos'], cell_format)
                    row += 1

            # Total del tipo
            worksheet.write(row, 0, f'Total {tipo}', label_format)
            worksheet.write(row, 3, datos['total_min'], total_format)
            row += 1

    row += 1

    # ABASTECIMIENTOS
    worksheet.merge_range(row, 0, row, 3, '*Abastecimientos*', section_format)
    row += 1

    worksheet.write(row, 0, 'Descripcion', header_format)
    worksheet.write(row, 1, 'Litros', header_format)
    worksheet.write(row, 2, 'Hora', header_format)
    worksheet.write(row, 3, 'Horometro', header_format)
    row += 1

    for abast, datos in abastecimientos_data.items():
        worksheet.write(row, 0, abast, label_format)
        worksheet.write(row, 1, datos['Litros'], cell_format)
        worksheet.write(row, 2, datos['Hora'], cell_format)
        worksheet.write(row, 3, datos['Horometro'], cell_format)
        row += 1

    row += 1

    # DECLARACION DE INCIDENTES
    worksheet.merge_range(row, 0, row, 3, '*Declaracion de Incidentes*', section_format)
    row += 1
    worksheet.write(row, 0, 'Incidentes Ambientales/Personas/Equipos', label_format)
    worksheet.merge_range(row, 1, row, 3, incidentes_tipo, value_format)
    row += 1
    worksheet.write(row, 0, 'Descripcion', label_format)
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

    # Insertar firmas como imagenes
    if firma_operador_file is not None:
        firma_operador_file.seek(0)
        img_data = BytesIO(firma_operador_file.read())
        worksheet.write(firma_row, 0, 'Firma Operador:', label_format)
        worksheet.insert_image(firma_row + 1, 0, 'firma_operador.png',
                             {'image_data': img_data, 'x_scale': 0.3, 'y_scale': 0.3})

    if firma_supervisor_rocmin_file is not None:
        firma_supervisor_rocmin_file.seek(0)
        img_data = BytesIO(firma_supervisor_rocmin_file.read())
        worksheet.write(firma_row, 2, 'Firma Supervisor ROCMIN:', label_format)
        worksheet.insert_image(firma_row + 1, 2, 'firma_supervisor_rocmin.png',
                             {'image_data': img_data, 'x_scale': 0.3, 'y_scale': 0.3})

    if firma_supervisor_cliente_file is not None:
        firma_supervisor_cliente_file.seek(0)
        img_data = BytesIO(firma_supervisor_cliente_file.read())
        worksheet.write(firma_row, 4, 'Firma Supervisor Cliente:', label_format)
        worksheet.insert_image(firma_row + 1, 4, 'firma_supervisor_cliente.png',
                             {'image_data': img_data, 'x_scale': 0.3, 'y_scale': 0.3})

    workbook.close()
    output.seek(0)
    return output

# ============================================
# BOTON DE DESCARGA
# ============================================
st.markdown("---")
st.header("Descargar Registro")

if st.button("Generar Excel para Descarga", type="primary"):
    if not operador:
        st.error("Por favor, ingrese el nombre del Operador.")
    else:
        # Generar nombre del archivo: fecha_Turno_NombreOperador
        fecha_str = fecha.strftime("%Y%m%d")
        turno_str = turno
        operador_str = operador.replace(" ", "_")
        nombre_archivo = f"{fecha_str}_{turno_str}_{operador_str}.xlsx"

        excel_file = generar_excel()

        st.download_button(
            label=f"Descargar: {nombre_archivo}",
            data=excel_file,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success(f"Excel generado correctamente. Total Metros Reales: {total_metros:.2f} m")

# ============================================
# FOOTER
# ============================================
st.markdown("---")
st.markdown("*ROCMIN - Sistema de Registro de Produccion Diaria*")
