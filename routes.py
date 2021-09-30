from flask import Blueprint, render_template, request, redirect, url_for
from werkzeug.utils import secure_filename
import os
import pyttsx3
import docx
import logging
import concurrent.futures
import zipfile
import multiprocessing
from win32com import client as wc

cota = Blueprint('cota', __name__)

#Rutas del servidor

# Ruta inicial
@cota.route("/")
def index():
    files = os.listdir(os.getcwd() + '/archivos')  
    for file in files:
        os.remove(os.getcwd() + '/archivos/' + file)
    
    audios = os.listdir(os.getcwd() + '/static/audios')
    for audio in audios:
        os.remove(os.getcwd() + '/static/audios/' + audio)

    filePath = r"" + os.getcwd() + "/static/audios-comprimidos.zip"
    if os.path.exists(filePath):
        os.remove(os.getcwd() + '/static/audios-comprimidos.zip')


    return render_template('index.html')

# Ruta para guardar los archivos en la carpeta del servidor
@cota.route('/upload', methods=['POST'])
def upload():
    if request.method == 'POST':
        fileText = request.files.getlist('textos') 
        for file in fileText:
            try:
                filename = secure_filename(file.filename)
                #print("Este es el nombre: ", filename)
                #newName = filename.split('_')
                #longitud = len(newName)
                file.save(os.getcwd() + "/archivos/" + filename)
            except FileNotFoundError:
                return 'Hay un error'
    return redirect(url_for('.converting'))

# Ruta para buscar un texto en archivos
@cota.route('/buscar', methods=['GET', 'POST'])
def buscar():
    if request.method == 'POST':
        textoBuscar = request.form['palabra']
        result = buscarContenido(textoBuscar.lower())
        return result
    else:
        return render_template('view-audios.html')

# Ruta para convertir los archivos de texto a audio             
@cota.route('/convert')
def converting():
    archivos = os.listdir(os.getcwd() + '/archivos')
    #print(os.cpu_count()-1)

    futures = []
    for i in archivos:
        name = i.split(".")
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=multiprocessing.cpu_count()-1) as executor:
            futures.append(executor.submit(convertingFiles, name, i))

    for future in futures:
        print(future.result())

    #Llamado a la funcion para comprimir la carpeta de los archivos de audio en .zip
    compress()

    return redirect(url_for('.view_audio'))

# Ruta para mostrar los archivos de audio en una vista html
@cota.route('/complete-conversion')
def view_audio():
    files = (os.listdir(os.getcwd() + '/static/audios'))
    longitud = "({})".format(len(files))
    return render_template('view-audios.html', audios = files, lon = longitud)

#Funciones requeridas

def buscarContenido(texto):
    listaDeArchivos = list()
    listaEncontrados = list()

    listaDeArchivos = (os.listdir(os.getcwd()+'/archivos'))
    futures = []
    for archivo in listaDeArchivos:
        name = archivo.split('.')

        with concurrent.futures.ThreadPoolExecutor(max_workers=multiprocessing.cpu_count()-1) as executor:
            futures.append(executor.submit(buscarContenidoArchivo, name, archivo, texto))

    for future in futures:
        if future.result() != False:
            listaEncontrados.append(future.result()) 

    return render_template('view-audios.html', audios = listaEncontrados)     

def buscarContenidoArchivo(name, archivo, texto):

    if name[1] == 'txt' or name[1] == 'dat':
        f = open(os.getcwd() + '/archivos/' + archivo, 'r', encoding='UTF-8')
        contenido = f.read().lower()
        
        if contenido.find(texto) != -1:
            nombre = name[0] + '.mp3'
            return nombre
        else:
            return False
    elif name[1] == 'docx':
        contenido = getTextDocx(archivo).lower()

        if contenido.find(texto) != -1:
            nombre = name[0] + '.mp3'
            return nombre
        else:
            return False
    elif name[1] == 'doc':
        word = wc.Dispatch('Word.Application')
        document = word.Documents.Open(os.getcwd() + '/archivos/' + archivo)
        document.SaveAs(os.getcwd() + '/archivos/' + name[0] + '.docx', 16)
        document.Close(False)
        os.remove(os.getcwd() + '/archivos/' + archivo)

        contenido = getTextDocx(name[0] + '.docx').lower()

        if contenido.find(texto) != -1:
            nombre = name[0] + '.mp3'
            return nombre
        else:
            return False

def convertTextToAudio(texto, namefile):
    engine = pyttsx3.init()
    engine.setProperty("rate", 140)

    engine.save_to_file(texto, os.getcwd() + '/static/audios/' + namefile + '.mp3')
    engine.runAndWait()

    return namefile + ' convertido'

def compress():
    audios = os.listdir(os.getcwd() + '/static/audios')

    with zipfile.ZipFile(os.getcwd() + '/static/audios-comprimidos.zip', 'w') as fzip:
        for audio in audios:
            fzip.write(os.getcwd() + '/static/audios/' + audio)
    
    print('Compression completed')       

#Funcion para obtener el texto en un archivo .docx
def getTextDocx(texto):
    doc = docx.Document(os.getcwd() + '/archivos/' + texto)
    fullText = []

    for parrafo in doc.paragraphs:
        fullText.append(parrafo.text)
    
    return '\n'.join(fullText)

def convertingFiles(name, i):

    if name[1] == "txt" or name[1] == "dat":
        f = open(os.getcwd() + '/archivos/' + i, 'r', encoding='UTF-8')
        texto = f.read()    
        name = i.split(".")

        convertTextToAudio(texto, name[0])

        return name[0] + ' convertido'

    elif name[1] == "docx":      
        texto = getTextDocx(i)

        convertTextToAudio(texto, name[0])

        return name[0] + ' convertido'

    elif name[1] == "doc":
        w = wc.Dispatch('Word.Application')
        doc = w.Documents.Open(os.getcwd() + '/archivos/' + i)
        doc.SaveAs(os.getcwd() + '/archivos/' + name[0] + '.docx', 16)
        doc.Close(False)
        os.remove(os.getcwd() + '/archivos/' + i)

        texto = getTextDocx(name[0] + '.docx')

        convertTextToAudio(texto, name[0])

        return name[0] + ' convertido'

    else:
        os.remove(os.getcwd() + '/archivos/' + i)

        return "Borrado " + name[0]

    

