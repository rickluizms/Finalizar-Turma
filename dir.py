import os
import zipfile 
import shutil

def CriarPasta():
    os.makedirs('./Documentos/Cartas/word')
    os.mkdir('./Documentos/Cartas/PDF')

    os.makedirs('./Documentos/Certificados/word')
    os.mkdir('./Documentos/Certificados/PDF')

    os.makedirs('./Documentos/Lista/word')
    os.mkdir('./Documentos/Lista/PDF')

    os.makedirs('./Documentos/Termos/word')
    os.mkdir('./Documentos/Termos/PDF')

def Zip():

    docs_zip = zipfile.ZipFile('./documentos.zip', 'w')

    for folder, subfolders, files in os.walk('./Documentos'):
        for file in files:
            if file.endswith('.pdf'):
                docs_zip.write(os.path.join(folder, file), os.path.relpath(os.path.join(folder,file), 
                './Documentos'), compress_type = zipfile.ZIP_DEFLATED)
                
            if file.endswith('.docx'):
                docs_zip.write(os.path.join(folder, file), os.path.relpath(os.path.join(folder,file), 
                './Documentos'), compress_type = zipfile.ZIP_DEFLATED)
    
    docs_zip.close()


def delete():
    shutil.rmtree('./Documentos')