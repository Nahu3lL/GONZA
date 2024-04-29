import os
import pandas as pd
import mysql.connector
import requests
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import random
import zipfile
import time
import shutil
import DBpassword

connection = mysql.connector.connect(
    host="localhost",
    user="root",
    password=DBpassword,
    database="mupema"
)
cursor = connection.cursor()

foldersCRT = []

timestampC = datetime.now().strftime("%Y-%m-%d--%H-%M-%S")
timestampS = datetime.now().strftime("%Y-%m-%d")

directoryIMG = 'barcodeImages'

# Espera archivos

file_path = 'uploaded_files.txt'

prev_modification_time = os.path.getmtime(file_path)
    
while True:
    time.sleep(1) # 1seg

    current_modification_time = os.path.getmtime(file_path)
        
    if current_modification_time != prev_modification_time:
        with open(file_path, 'r') as file:
            file_contents = file.read()

        print(file_contents)
        prev_modification_time = current_modification_time
        break

df = pd.read_excel(str(file_contents))


def generate_unique_codes(existing_codes, num_codes):
    generated_codes = set()
    while len(generated_codes) < num_codes:
        new_code = str(random.randint(10**12, 10**13 - 1))
        if new_code not in existing_codes and new_code not in generated_codes:
            generated_codes.add(new_code)
    return list(generated_codes)

#m.py

def crear_pdf_con_cuadrado(PDFConfig, titulo, presentacion, img_name, container, almacen, pallet, caja, item, unidad):
    img_path = "C:\\Users\\NADMIN\\Desktop\\barcodeImages\\"

    PDFConfig = os.path.join(fullPathD, PDFName)

    ancho_hoja, alto_hoja = letter
    c = canvas.Canvas(PDFConfig, pagesize=letter)

    # Primera sección
    ancho_cuadrado = (ancho_hoja - 40)
    alto_cuadrado = (alto_hoja - 40) / 2

    x_ancho_cuadrado = 20
    y_alto_cuadrado = alto_hoja - alto_cuadrado - 20

    c.rect(x_ancho_cuadrado, y_alto_cuadrado, ancho_cuadrado, alto_cuadrado)

    ancho_titulo = ancho_cuadrado
    alto_titulo = alto_cuadrado / 3

    x_ancho_titulo = 20
    y_alto_titulo = 20 + (alto_cuadrado * 2) - alto_titulo

    c.rect(x_ancho_titulo, y_alto_titulo, ancho_titulo, alto_titulo)
    c.setFont("Helvetica", 50)
    ancho_texto_titulo = c.stringWidth(titulo)
    x_texto_titulo = x_ancho_titulo + (ancho_titulo - ancho_texto_titulo) / 2
    y_texto_titulo = (y_alto_titulo + (alto_titulo - c._leading) / 2) + 30
    c.drawString(x_texto_titulo, y_texto_titulo, titulo)

    ancho_cb = ancho_cuadrado * (2 / 3)
    alto_cb = alto_cuadrado * (2 / 3)

    x_ancho_cb = 20
    y_alto_cb = alto_hoja - alto_cuadrado - 20

    c.rect(x_ancho_cb, y_alto_cb, ancho_cb, alto_cb)
    for filename in img_name:
        imagen_cb = img_path + filename
        x_img_cb = x_ancho_cb + (ancho_cb - 250) / 2
        y_img_cb = y_alto_cb + (alto_cb - 175) / 2
        c.drawImage(imagen_cb, x_img_cb, y_img_cb, width=250, height=175)

    ancho_container = ancho_cuadrado / 3
    alto_container = alto_cuadrado * (2/9)

    x_ancho_container = ancho_cb + 20
    y_alto_container = alto_hoja - alto_cuadrado - 20

    c.rect(x_ancho_container, y_alto_container, ancho_container, alto_container)
    c.setFont("Helvetica", 100)
    ancho_texto_container = c.stringWidth(container)
    x_texto_container = x_ancho_container + (ancho_container - ancho_texto_container) / 2
    y_texto_container = (y_alto_container + (alto_container - c._leading) / 2) + 25
    c.drawString(x_texto_container, y_texto_container, container)

    ancho_pallet = ancho_cuadrado / 3
    alto_pallet = alto_cuadrado * (2/9)

    x_ancho_pallet = ancho_cb + 20
    y_alto_pallet = alto_hoja - alto_cuadrado - 20 + alto_container

    c.rect(x_ancho_pallet, y_alto_pallet, ancho_pallet, alto_pallet)
    c.setFont("Helvetica", 100)
    ancho_texto_pallet = c.stringWidth(pallet)
    x_texto_pallet = x_ancho_pallet + (ancho_pallet - ancho_texto_pallet) / 2
    y_texto_pallet = (y_alto_pallet + (alto_pallet - c._leading) / 2) + 25
    c.drawString(x_texto_pallet, y_texto_pallet, pallet)

    ancho_item = ancho_cuadrado / 3
    alto_item = alto_cuadrado * (2/9)

    x_ancho_item = ancho_cb + 20
    y_alto_item = alto_hoja - alto_cuadrado - 20 + (alto_pallet * 2)

    c.rect(x_ancho_item, y_alto_item, ancho_item, alto_item)
    c.setFont("Helvetica", 100)
    ancho_texto_item = c.stringWidth(item)
    x_texto_item = x_ancho_item + (ancho_item - ancho_texto_item) / 2
    y_texto_item = (y_alto_item + (alto_item - c._leading) / 2) + 25
    c.drawString(x_texto_item, y_texto_item, item)

    # Segunda sección
    ancho_rectangulo = (ancho_hoja - 40)
    alto_rectangulo = (alto_hoja - 40) / 2

    x_ancho_rectangulo = 20
    y_alto_rectangulo = 20

    c.rect(x_ancho_rectangulo, y_alto_rectangulo, ancho_rectangulo, alto_rectangulo)

    ancho_presentacion = ancho_rectangulo
    alto_presentacion = alto_rectangulo / 3

    x_ancho_presentacion = 20
    y_alto_presentacion = 20 + alto_rectangulo - alto_presentacion

    c.rect(x_ancho_presentacion, y_alto_presentacion, ancho_presentacion, alto_presentacion)
    c.setFont("Helvetica", 50)
    ancho_texto_presentacion = c.stringWidth(presentacion)
    x_texto_presentacion = x_ancho_presentacion + (ancho_presentacion - ancho_texto_presentacion) / 2
    y_texto_presentacion = (y_alto_presentacion + (alto_presentacion - c._leading) / 2) + 30
    c.drawString(x_texto_presentacion, y_texto_presentacion, presentacion)

    ancho_ad = ancho_rectangulo * (2 / 3)
    alto_ad = alto_rectangulo * (2 / 3)

    x_ancho_ad = 20
    y_alto_ad = 20

    c.rect(x_ancho_ad, y_alto_ad, ancho_ad, alto_ad)
    for filename in img_name:
        imagen_ad = img_path + filename
        x_img_ad = x_ancho_ad + (ancho_ad - 250) / 2
        y_img_ad = y_alto_ad + (alto_ad - 175) / 2
        c.drawImage(imagen_ad, x_img_ad, y_img_ad, width=250, height=175)

    ancho_almacen = ancho_rectangulo / 3
    alto_almacen = alto_rectangulo * (2 / 9)

    x_ancho_almacen = ancho_ad + 20
    y_alto_almacen = 20

    c.rect(x_ancho_almacen, y_alto_almacen, ancho_almacen, alto_almacen)
    c.setFont("Helvetica", 100)
    ancho_texto_almacen = c.stringWidth(almacen)
    x_texto_almacen = x_ancho_almacen + (ancho_almacen - ancho_texto_almacen) / 2
    y_texto_almacen = (y_alto_almacen + (alto_almacen - c._leading) / 2) + 25
    c.drawString(x_texto_almacen, y_texto_almacen, almacen)

    ancho_caja = ancho_rectangulo / 3
    alto_caja = alto_rectangulo * (2 / 9)

    x_ancho_caja = ancho_ad + 20
    y_alto_caja = 20 + alto_almacen

    c.rect(x_ancho_caja, y_alto_caja, ancho_caja, alto_caja)
    c.setFont("Helvetica", 100)
    ancho_texto_caja = c.stringWidth(caja)
    x_texto_caja = x_ancho_caja + (ancho_caja - ancho_texto_caja) / 2
    y_texto_caja = (y_alto_caja + (alto_caja - c._leading) / 2) + 25
    c.drawString(x_texto_caja, y_texto_caja, caja)

    ancho_unidad = ancho_rectangulo / 3
    alto_unidad = alto_rectangulo * (2 / 9)

    x_ancho_unidad = ancho_ad + 20
    y_alto_unidad = 20 + alto_almacen + alto_caja

    c.rect(x_ancho_unidad, y_alto_unidad, ancho_unidad, alto_unidad)
    c.setFont("Helvetica", 100)
    ancho_texto_unidad = c.stringWidth(unidad)
    x_texto_unidad = x_ancho_unidad + (ancho_unidad - ancho_texto_unidad) / 2
    y_texto_unidad = (y_alto_unidad + (alto_unidad - c._leading) / 2) + 25
    c.drawString(x_texto_unidad, y_texto_unidad, unidad)

    c.save()

#qd.py

def download_barcode_images(data_list, directoryIMG):
    success_count = 0
    error_messages = []
    downloaded_files = []

    os.makedirs(directoryIMG, exist_ok=True)

    for data in data_list:
        try:
            url = f"https://barcode.tec-it.com/barcode.ashx?data={data}"
            response = requests.get(url)
            if response.status_code == 200:
                filename = f"{data}.png"
                filepath = os.path.join(directoryIMG, filename)
                with open(filepath, "wb") as f:
                    f.write(response.content)
                success_count += 1
                downloaded_files.append(filename)
            else:
                error_messages.append(f"Failed to download image for data: {data}")
        except Exception as e:
            error_messages.append(f"Error downloading image for data {data}: {str(e)}")
    
    return success_count, error_messages, downloaded_files

#c.py

required_columns = ['propietario_cuit', 'codigo', 'descripcion', 'lote', 'bulto1', 'bulto1_cantidad', 'bulto1_kg', 'bulto1_m3', 'bulto1_valorDeclarado1', 'bulto2', 'bulto2_cantidad', 'bulto2_kg', 'bulto2_m3', 'bulto3', 'bulto3_cantidad', 'bulto3_kg', 'bulto3_m3', 'serie', 'idexterno', 'transportista_cuit', 'CONT', 'PALL', 'UN', 'STATUS', 'BCOD', 'FENTRADA']
required_completed_columns = ['propietario_cuit', 'CONT', 'PALL', 'UN', 'STATUS']
missing_columns = [col for col in required_columns if col not in df.columns]
existing_codes = set(df['BCOD'])

if missing_columns:
    print(f"Faltan las siguientes columnas: {', '.join(missing_columns)}")
else:
    if 'STATUS' in df.columns and (df['STATUS'] == 'S').any():
        rows_with_s = df[df['STATUS'] == 'S']   # Genera valores random para articulos habilitados ("S")
        barcodes = [int(random.randint(10**12, 10**13 - 1)) for _ in range(len(rows_with_s))]
        df.loc[df['STATUS'] == 'S', 'BCOD'] = barcodes

        preECodes = "SELECT BCOD FROM register WHERE STATUS = %s"
        cursor.execute(preECodes, ('S',))
        repitedCodes = cursor.fetchall()
        activeBC = []
        for row in repitedCodes:
            if 'BCOD' in row:
                activeBC.append(row['BCOD'])
            else:
                activeBC.append(0)


        while True:
            if any(code in activeBC or code in df['BCOD'] for code in barcodes):
                new_barcodes = generate_unique_codes(existing_codes, len(existing_codes))
                #print("Codigos recreados: " + new_barcodes)
            
                replacement_values = dict(zip(df.loc[df['BCOD'].isin(existing_codes), 'BCOD'], new_barcodes))
                df['BCOD'].replace(replacement_values, inplace=True)
            else:
                break

        df.loc[df['STATUS'] == 'N', 'BCOD'] = 0   # Da valores de 0 para los articulos no habilitados ("N")
            
        df.to_excel(file_contents, index=False)

        BCOD = df['BCOD'].tolist()
        data_list = [str(code) for code in BCOD]
        success_count, error_messages, downloaded_files = download_barcode_images(data_list, directoryIMG)
        if error_messages:
            print("Errors occurred during image download:")
            for error in error_messages:
                print(error)
    else:
        print("No rows with 'STATUS' of 'S' found.")


CONT_list = []
PALL_list = []
UN_list = []
COD_list = []
propietario_cuit_list = []
descripcion_list = []
downloaded_filenames = []


for index, row in df.iterrows():
    propietario_cuit = row['propietario_cuit']
    descripcion = row['descripcion']
    CONTR = str(row['CONT'])
    CONT = 'C' + str(row['CONT'])
    PALLR = str(row['PALL'])
    PALL = 'P' + str(row['PALL'])
    UNR = str(row['UN'])
    UN = 'I' + str(row['UN'])
    BCOD = row['BCOD']
    
    cursor = connection.cursor()
    cursor.execute("SELECT NCLIENT FROM client WHERE CUIT = %s", (propietario_cuit,))
    clientE = cursor.fetchone()

    if clientE:
        titulo = clientE[0]
    else:
        print("Cliente no asignado")
        titulo = propietario_cuit

    success_count, error_messages, downloaded_files = download_barcode_images([BCOD], 'barcodeImages')

    downloaded_filenames.extend(downloaded_files)

    # PArametros del PDF
    img_name = downloaded_files
    container = CONT
    pallet = PALL
    item = UN
    presentacion = titulo
    almacen = container
    caja = pallet
    unidad = item
    counter = 1
    while True:
        directoryPDF = f"Op.{counter}-{titulo}-{timestampS}"
        fullPathD = os.path.join("C:\\Users\\NADMIN\\Desktop\\OPERACIONES", directoryPDF)
        if not os.path.exists(fullPathD):
            os.makedirs(fullPathD)
            foldersCRT.append(fullPathD)
            break
        counter += 1

    PDFName = f"{titulo}-{str(BCOD)}.pdf"
    PDFConfig = os.path.join(fullPathD, PDFName)

    crear_pdf_con_cuadrado(PDFConfig, titulo, presentacion, img_name, container, almacen, pallet, caja, item, unidad)

for index, row in df.iterrows():
    queryIndexTable = "INSERT INTO register (CUIT, CODE, DESCRIPTION, LOTE, BULTO1, BULTO1CANT, BULTO1Kg, BULTO1M3, BULTO1VD, BULTO2, BULTO2CANT, BULTO2Kg, BULTO2M3, BULTO3, BULTO3CANT, BULTO3Kg, BULTO3M3, SERIE, EXTID, TCUIT, CONT, PALL, UN, STATUS, BCOD, FENTRADA) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
    
    values = (
        row['propietario_cuit'], row['codigo'], row['descripcion'], row['lote'],
        row['bulto1'], row['bulto1_cantidad'], row['bulto1_kg'], row['bulto1_m3'], row['bulto1_valorDeclarado1'],
        row['bulto2'], row['bulto2_cantidad'], row['bulto2_kg'], row['bulto2_m3'],
        row['bulto3'], row['bulto3_cantidad'], row['bulto3_kg'], row['bulto3_m3'],
        row['serie'], row['idexterno'], row['transportista_cuit'],
        row['CONT'], row['PALL'], row['UN'], row['STATUS'], row['BCOD'],
        timestampC
    )
    cursor.execute(queryIndexTable, values)

connection.commit()
cursor.close()
connection.close()

zipName = f"Op.from{timestampC}.zip"
zip_file_name = os.path.join("C:\\Users\\NADMIN\\Desktop\\OPEXPO", zipName)

with zipfile.ZipFile(zip_file_name, 'w') as zipf:
    for folder in foldersCRT:
        for root, dirs, files in os.walk(folder):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), folder))
