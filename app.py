import streamlit as st
from io import BytesIO
from pathlib import Path
from openpyxl import load_workbook
from datetime import time, datetime
from urllib.parse import urlparse
from zipfile import BadZipFile


DIAS = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo']
CAMPOS_PLAN = ['dia', 'hora_inicio','hora_fin','tiempo_comida','menu', 'receta_id']
CAMPOS_RECETAS = ['receta_id','nombre','ingredientes_cantidades','preparacion',
                  'video_url','estado','fuente_preparacion']
BASE = Path(__file__).resolve().parent

def texto(valor):
    ##Si una celda vacia se repesenta como texto vacio, nunca como None.
    return '' if valor is None else str(valor).strip()

def leer_hoja(libro,nombre,campos):
    if nombre not in libro.sheetnames:
        raise ValueError(f'Falta la hoja {nombre}.')
    filas = list(libro[nombre].iter_rows(values_only=True))
    if not filas:
        raise ValueError(f'La hoja {nombre} está vacía.')
    encabezados = [texto(celda) for celda in filas[0]]
    if any(encabezados.count(campo)!=1 for campo in campos):
        raise ValueError(f'Revisa los encabezados de {nombre}: {', '.join(campos)}.')
    registros = []

    for numero, fila in enumerate(filas[1:], start=2):
        if not any(texto(celda) for celda in fila):
            continue
        if any(isinstance(celda, str) and celda.startswith('=') for celda in fila):
            raise ValueError(f'{nombre}, fila{numero}: usa valores, no formulas.')
        registro = {campo:fila[encabezados.index(campo)] for campo in campos}
        registros.append(registro)
    if not registros:
        raise ValueError(f'La hoja {nombre} no contiene registros.')
    return registros

def cargar_plan(contenido):
    libro = load_workbook(BytesIO(contenido), read_only=True, data_only=False)
    try:
        plan = leer_hoja(libro,'Plan',CAMPOS_PLAN)
        filas_recetas = leer_hoja(libro,"Recetas", CAMPOS_RECETAS)
    finally:
        libro.close()
    
    recetas = {}
    for receta in filas_recetas:
        receta ={campo: texto(valor) for campo,valor in receta.items()}
        clave = receta['receta_id']
        
        if not clave or clave in recetas:
            raise ValueError(f'Identificador de receta vacío o repetido: {clave!r}.')
        if not receta['nombre'] or not receta['ingredientes_cantidades']:
            raise ValueError(f'{clave}: faltan nombre o alimentos y cantidades')
        if receta['video_url'] and not enlace_valido(receta['video_url']):
            raise ValueError(f'{clave}: video_url debe ser un enlace https completo')
        recetas[clave] = receta
    claves = set()
    for fila in plan:
        for campo in ['dia','tiempo_comida','menu','receta_id']:
            fila[campo]=texto(fila[campo])
            if not fila[campo]:
                raise ValueError(f'Plan: falta un valor {campo}')
        if fila['dia'] not in DIAS:
                raise ValueError(f'Día desconocido: {fila['dia']}')
        fila['inicio'] = minutos(fila['hora_inicio'])
        fila['fin'] = minutos(fila['hora_fin'])
        if fila['fin']< fila['inicio']:
            raise ValueError(f'el fin no puede ser anterior al inicio; nose admiten intervalos nocturnos')
        if fila['receta_id'] not in recetas:
            raise ValueError(f'No existe la receta {fila['receta_id']}')
        clave = (fila['dia'], fila['tiempo_comida'])
        if clave in claves:
            raise ValueError(f'Horario Duplicado {clave[0]}, {clave[1]}')
        claves.add(clave)
             
    return plan, recetas

def consultar_dia(plan,dia):
    filas =[fila for fila in plan if fila['dia']==dia]
    return sorted(filas,key=lambda fila: fila['inicio'])

def mostrar_menu(fila,recetas,minuto):
    receta = recetas[fila['receta_id']]
    inicio, fin = hora_visible(fila['hora_inicio']), hora_visible(fila['hora_fin'])
    horario = inicio if inicio == fin else f'{inicio}-{fin}'
    with st.container(border=True):
        st.caption(f'{horario} - {fila['tiempo_comida']} - {estado_horario(fila,minuto)} ')
        st.subheader(fila['menu'])
        st.write(receta['ingredientes_cantidades'])

        with st.expander('Preparacion y tutorial'):
            completa = all(receta[c] for c in ['preparacion', 'video_url','fuente_preparacion'])
            if receta['estado']=='Revisada' and completa:
                st.text(receta['preparacion'])
                st.link_button('Abrir tutorial de cocina',receta['video_url'])
                st.caption('Fuente de preparacion: '+receta['fuente_preparacion'])
            else:
                st.info('Preparacion y video pendientes de revision.')
    


def minutos(valor):
    ##Acepta HH:MM de la plantilla y horas nativas guardadas por Excel
    if isinstance(valor,time):
        return valor.hour*60 + valor.minute
    #14:30 = 14*60 + 30 = 87 minutos
    cadena = texto(valor)
    partes = cadena.split(':')
    #8:30 = ['8','30']
    #24:00 6:00
    if len(partes)!=2 or len(cadena)!= 5 or not all(p.isdigit() for p in partes):
        raise ValueError(f'Horas inválidas: {cadena!r}. Usa HH:MM, por ejemplo 14:00')
    hora ,minuto = int(partes[0]), int(partes[1])
    if not 0 <= hora <= 23 or not 0<= minuto <= 59:
        raise ValueError(f'Revisa la hora. => {cadena} fuera de formato.')
    return hora*60 + minuto

def hora_visible(valor):
    total = minutos(valor)
    return f'{total//60:02d}:{total%60:02d}'

def enlace_valido(url):
    partes = urlparse(texto(url))
    return partes.scheme == 'https' and bool(partes.netloc)

def estado_horario(fila,minuto):
    if minuto < fila['inicio']:
        return 'Proximo'
    if minuto <= fila['fin']:
        return 'En Horario'
    return 'Horario Transcurrido'




        
        





def main():
    st.set_page_config(page_title='Mi menú de hoy - Ruben Rivera', page_icon='🍽️', layout='centered')
    st.title('Mi menú de hoy - Ruben Rivera')
    st.write('Consulta los horarios , alimento, cantidades de tu documento')
    archivo = st.file_uploader('Cargar mi Excel', type=['xlsx'])
    try:
        contenido = archivo.getvalue() if archivo else (BASE/'Mi_menu_de_hoy.xlsx').read_bytes()
        plan, recetas = cargar_plan(contenido)
    except(ValueError, KeyError,OSError, BadZipFile) as error:
        st.error(f'No pude cargar el plan: {error}')
        st.stop()
    st.caption('Archivo cargado' + (archivo.name if archivo else 'Mi_menu_de_hoy.xlsx'))
    modo = st.radio('Consulta por',['Dia de la semana','Dia del ciclo de 15 Dias'], horizontal=True)
    if modo == 'Dia de la semana':
        dia = st.selectbox('Día', DIAS)
    else:
        numero = st.number_input('Dia del ciclo',min_value=1, max_value=15, value=1,step=1)
        dia = DIAS[(numero-1)%7]
        st.caption(f'Dia {numero}: {dia}. El ciclo documentado comienza en lunes.')
    ahora = datetime.now()
    hora = st.time_input('Hora consultada', value=ahora.strftime('%H:%M'), step=60)
    #hora = st.time_input('Hora consultada', value=time(14,30), step=60)

    st.caption('La hora es manual para probar la agenda; no es un reloj ni envia avisos')
    minuto = hora.hour *60 + hora.minute


    filas = consultar_dia(plan, dia)
    solo_actual = st.checkbox('Mostrar solo lo que corresponde a esta hora')
    actuales = [fila for fila in filas if estado_horario(fila,minuto)== 'En Horario']
    if not filas:
        st.warning('No hay registros para el dia seleccionado.')
    elif not actuales:
        siguientes = [fila for fila in filas if fila['inicio']>minuto]
        if siguientes:
            siguiente = siguientes[0]
            st.info(f'Fuera de los horarios indicados. Siguiente: {siguiente['tiempo_comida']} a las {hora_visible(siguiente['hora_inicio'])}.')
                                                                                                        #hora_visible['hora_inicio']
        else:
            st.info('Los horarios indicados para este dia ya transcurrieron. Puedes consultar otro dia.')
    for fila in actuales if solo_actual else filas:
        mostrar_menu(fila, recetas, minuto)

    with st.expander('Sobre este plan'):
        st.write('Distribucion tomada de Horario Semanal')

   

if __name__ == '__main__':
    main()