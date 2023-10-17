from datetime import datetime, timedelta
from docxtpl import DocxTemplate
import pandas as pd
import streamlit as st
import zipfile

def generar_data_cronograma(inversion:int, meses:int, tasa_interes:int, fecha_inicio:str):
    """
    Genera un cronograma de pagos de acuerdo a los parametros de entrada
    :param inversion: Monto de la inversion
    :param meses: Numero de meses del cronograma
    :param tasa_interes: Tasa de interes mensual
    :param fecha_inicio: Fecha de inicio del cronograma
    :return: Cronograma de pagos
    """
    cronogram_data = []
    interes_total = inversion*tasa_interes*meses/12
    monto_total = inversion + interes_total
    interes_mensual = interes_total/meses

    for cuota_index in range(1,meses+1):
        cuota_data = {}
        cuota_data['cuota'] = cuota_index
        cuota_data['fecha'] = (fecha_inicio + timedelta(days=35)*(cuota_index-1)).strftime("%d/%m/%Y")
        cuota_data['capital'] = inversion
        cuota_data['interes'] = interes_mensual
        cronogram_data.append(cuota_data)

    return cronogram_data, monto_total

st.markdown('# Generación de contratos')
st.markdown('_Esta herramienta permite generar los contratos de manera dinámica y masiva._')

# uploads
upload_template_word = st.file_uploader('Paso 1 : Sube el archivo docx con el template', type=['docx'])
upload_file = st.file_uploader('Paso 2: Sube el archivo excel para generar los contratos', type=['xlsx'])

# check if file are uploaded
if upload_template_word is not None:
    template = DocxTemplate(upload_template_word)

if upload_file is not None:
    df_list = pd.read_excel(upload_file)
    st.dataframe(df_list.head(10)) # show first 10

if st.button('Generar contratos'):
    list_docs = []

    for index, row in df_list.iterrows():
        nombre = row['nombre']
        inversion = row['inversion']
        meses = row['meses']
        tasa_interes = row['tasa_interes']
        fecha_inicio = row['fecha_inicio'].to_pydatetime()
        nro_dni = row['nro_dni']
        estado_civil = row['estado_civil']
        nro_celular = row['nro_celular']
        direccion = row['direccion']
        nacionalidad = row['nacionalidad']

        cronograma_data, monto_total = generar_data_cronograma(inversion, meses, tasa_interes, fecha_inicio)
        context = {
        'nombre': nombre,
        'nro_dni': nro_dni,
        'estado_civil': estado_civil,
        'nro_celular': nro_celular,
        'nacionalidad':nacionalidad,
        'direccion':direccion,
        'inversion': inversion,
        'cuotas': cronograma_data,
        'monto_total': monto_total
        }

        file_name = f"contrato-{nombre}-{str(datetime.now())}.docx"

        template.render(context)
        template.save(file_name)
        list_docs.append(file_name)

    st.success('Se generaron los contratos satisfactoriamente')
    
    zip_name = f"contratos-{str(datetime.now())}.zip"
    with zipfile.ZipFile(zip_name, "w") as zipF:
        for invoice in list_docs:
            zipF.write(invoice, compress_type=zipfile.ZIP_DEFLATED)

    with open(zip_name, "rb") as fp:
        st.download_button(
        label="Download ZIP",
        data=fp,
        file_name="contratos.zip",
        mime="application/zip"
        )
       