import PyPDF2
import xlsxwriter
import re
import glob
import os
import sys
from os import listdir

def contadorAlumnosReprobadosTotales():
    #obtiene una ruta a partir de un filechooser
    workbook = xlsxwriter.Workbook('uploads/resultado.xlsx')
    worksheet = workbook.add_worksheet()
    directory = os.path.join('uploads/')

    test = os.listdir(directory)
    for item in test:
        if item.endswith(".xlsx"):
            os.remove(os.path.join(directory, item))


    #filename = askopenfilenames()
    #files = list(filename)
    files = glob.glob(directory+"*.pdf")
    profesores = list()
    aprobadosLista = list()
    reprobadosLista = list()
    worksheet.set_column('A:A', 40)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 20)
    worksheet.set_column('D:D', 20)
    worksheet.set_column('E:E', 20)

    for x in range(0,len(files)):
    # Crea un objeto pdf con la ruta obtenida
        pdfFileObj = open(files[x], 'rb')

    # Crea un objeto del tipo pdf reader
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

    # Imprime el numero de paginas del archivo
        #print(pdfReader.numPages)
        j = 0;
        aprobados = 0;
        reprobados = 0;
    #mientras j sea menor al numero de paginas del pdf
        while (j<pdfReader.numPages):

        # Crea un objeto tipo pagina
            pageObj = pdfReader.getPage(j)

        # extracting text from page
        # print(pageObj.extractText())

        # extra todo el texto de la pagina
            texto = pageObj.extractText();
            sinDerecho = texto.count("SIN DERECHO");
            noPresento = texto.count("NO PRESENTó EXAMEN");
            noPresento = noPresento + texto.count("NO PRESENTO EXAMEN");
        # obtiene todos los numeros contenidos en el pdf
            numeros = re.findall('\d+', texto)
            reprobados = reprobados + sinDerecho + noPresento;
           
            i = 12
        # mientras i sea menor a la cantidad de numeros, busca los numeros menores o iguales a 100
        # y los mayores o iguales a 60 y suma uno a la cantidad de aprobados en caso de encontrarlo
            while (i < len(numeros) - 2):
                if (int(numeros[i]) >= 60 and int(numeros[i]) <= 100):
                    print(numeros[i])
                    aprobados = aprobados + 1
                elif (int(numeros[i]) >= 0 and int(numeros[i]) < 60):
                    print(numeros[i])
                    reprobados = reprobados + 1
                i = i + 1
            j = j + 1




    #imprime la cantidad de aprobados
        result = re.search('ESCOLARDMA(.*)Maestro', texto)
        print(result.group(1))
        print("Aprobados:")
        print(aprobados)
        print("Reprobados:")
        print(reprobados)

        try:
            posicion = profesores.index(result.group(1))
            aprobadosLista[posicion] = aprobadosLista[posicion] + aprobados
            reprobadosLista[posicion] = reprobadosLista[posicion] + reprobados
        except ValueError:
            profesores.append(result.group(1))
            aprobadosLista.append(aprobados)
            reprobadosLista.append(reprobados)
    # se cierra el objeto pdf
        pdfFileObj.close()

    print(profesores)
    print(aprobadosLista)
    print(reprobadosLista)

    cell_format1 = workbook.add_format()
    cell_format1.set_num_format('0%')
    worksheet.write('A1', 'Profesor')
    worksheet.write('B1', 'Aprobados')
    worksheet.write('C1', 'Reprobados')
    worksheet.write('D1', '% de Aprobados')
    worksheet.write('E1', '% de Reprobados')

    pos = 2
    for z in range(0,len(profesores)):
        worksheet.write('D'+str(pos), '=B'+str(pos)+'/(B'+str(pos)+'+C'+str(pos)+')', cell_format1)
        worksheet.write('E'+str(pos), '=C'+str(pos)+'/(B'+str(pos)+'+C'+str(pos)+')', cell_format1)
        pos = pos + 1

    for z in range(0,len(profesores)):
        worksheet.write(z+1, 0, profesores[z])
        worksheet.write(z+1, 1, aprobadosLista[z])
        worksheet.write(z+1, 2, reprobadosLista[z])

    workbook.close()
    
    for item in test:
        if item.endswith(".pdf"):
            os.remove(os.path.join(directory, item))


def contadorAlumnosReprobadosNPSD():
    #obtiene una ruta a partir de un filechooser
    workbook = xlsxwriter.Workbook('uploads/resultado.xlsx')
    worksheet = workbook.add_worksheet()
    directory = os.path.join('uploads/')

    test = os.listdir(directory)
    for item in test:
        if item.endswith(".xlsx"):
            os.remove(os.path.join(directory, item))


    #filename = askopenfilenames()
    #files = list(filename)
    files = glob.glob(directory+"*.pdf")
    profesores = list()
    aprobadosLista = list()
    reprobadosLista = list()
    worksheet.set_column('A:A', 40)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 20)
    worksheet.set_column('D:D', 20)
    worksheet.set_column('E:E', 20)

    for x in range(0,len(files)):
    # Crea un objeto pdf con la ruta obtenida
        pdfFileObj = open(files[x], 'rb')

    # Crea un objeto del tipo pdf reader
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

    # Imprime el numero de paginas del archivo
        #print(pdfReader.numPages)
        j = 0;
        aprobados = 0;
        reprobados = 0;
    #mientras j sea menor al numero de paginas del pdf
        while (j<pdfReader.numPages):

        # Crea un objeto tipo pagina
            pageObj = pdfReader.getPage(j)

        # extracting text from page
        # print(pageObj.extractText())

        # extra todo el texto de la pagina
            texto = pageObj.extractText();
            sinDerecho = texto.count("SIN DERECHO");
            noPresento = texto.count("NO PRESENTó EXAMEN");
            noPresento = noPresento + texto.count("NO PRESENTO EXAMEN");
        # obtiene todos los numeros contenidos en el pdf
            numeros = re.findall('\d+', texto)
            reprobados = reprobados + sinDerecho + noPresento;

            i = 12
        # mientras i sea menor a la cantidad de numeros, busca los numeros menores o iguales a 100
        # y los mayores o iguales a 60 y suma uno a la cantidad de aprobados en caso de encontrarlo
            while (i < len(numeros) - 2):
                if (int(numeros[i]) >= 60 and int(numeros[i]) <= 100):
                    print(numeros[i])
                    
                elif (int(numeros[i]) >= 0 and int(numeros[i]) < 60):
                    print(numeros[i])
                    
                i = i + 1
            j = j + 1




    #imprime la cantidad de aprobados
        result = re.search('ESCOLARDMA(.*)Maestro', texto)
        print(result.group(1))
        print("Aprobados:")
        print(aprobados)
        print("Reprobados:")
        print(reprobados)

        try:
            posicion = profesores.index(result.group(1))
            aprobadosLista[posicion] = aprobadosLista[posicion] + aprobados
            reprobadosLista[posicion] = reprobadosLista[posicion] + reprobados
        except ValueError:
            profesores.append(result.group(1))
            aprobadosLista.append(aprobados)
            reprobadosLista.append(reprobados)
    # se cierra el objeto pdf
        pdfFileObj.close()

    print(profesores)
    print(aprobadosLista)
    print(reprobadosLista)

    cell_format1 = workbook.add_format()
    cell_format1.set_num_format('0%')
    worksheet.write('A1', 'Profesor')
    worksheet.write('B1', 'Reprobados')


    for z in range(0,len(profesores)):
        worksheet.write(z+1, 0, profesores[z])
        worksheet.write(z+1, 1, reprobadosLista[z])

    workbook.close()
    
    for item in test:
        if item.endswith(".pdf"):
            os.remove(os.path.join(directory, item))