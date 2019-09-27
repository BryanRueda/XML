#Importacion de librerias
import sys
import os
import xlsxwriter  

def readXML(path,date):

    #path='C:/Users/Bryan/Documents/UiPath/readXML'
    elements=[]

    #//////////////////////////////////////////////////////
    #Lectura de XML's existentes
    #//////////////////////////////////////////////////////

    #path = 'C:/Users/Bryan/Desktop/Lector de XML'

    files = []
    # r=root, d=directories, f = files
    for r, d, f in os.walk(path):
        for file in f:
            if '.xml' in file:
                files.append(os.path.join(r, file))
                

    #Bucle para cada uno de los xml encontrados
    for f in files:


        #Inicializacion de variables
        NoId=0
        AgenteRetencion=0
        TipoComprobante=0
        Establecimiento=0
        PuntoEmision=0
        Secuencial=0
        ClaveAcceso=0
        FechaEmision=0
        
        TipoComprobSus1=0
        EstablecimientoSus1=0
        PtoEmisionSus1=0
        SecuencialSus1=0
        FechaEmisionSus1=0
        CodigoRetencion1=0
        BaseImponible1=0
        PorcentajeRetencion1=0
        PorcentajeRetencion1=0
        ValorRetenido1=0
        TipoComprobSus2=0
        EstablecimientoSus2=0
        PtoEmisionSus2=0
        SecuencialSus2=0
        FechaEmisionSus2=0
        CodigoRetencion2=0
        BaseImponible2=0
        PorcentajeRetencion2=0
        ValorRetenido2=0
        TipoComprobSus3=0
        EstablecimientoSus3=0
        PtoEmisionSus3=0
        SecuencialSus3=0
        FechaEmisionSus3=0
        CodigoRetencion3=0
        BaseImponible3=0
        PorcentajeRetencion3=0
        ValorRetenido3=0
        TipoComprobSus4=0
        EstablecimientoSus4=0
        PtoEmisionSus4=0
        SecuencialSus4=0
        FechaEmisionSus4=0
        CodigoRetencion4=0
        BaseImponible4=0
        PorcentajeRetencion4=0
        ValorRetenido4=0
        
        
        Codigo1=0
        numDocSustento1=0
        Codigo2=0
        numDocSustento2=0
        Codigo3=0
        numDocSustento3=0
        Codigo4=0
        numDocSustento4=0

        TEXTO=0
        contador=0

#-------------------------------------------------------------------------------
# Extracción de INFO
#-------------------------------------------------------------------------------

        try:

            #Lectura de archivo XML inicial
            mytree = ET.parse(f)
            myroot = mytree.getroot()

            #Extraer una cadena de texto con toda la informacion
            for x in myroot.findall('comprobante'):
                TEXTO=(x.text)
                
            #Extraccion de parametros uno a uno
            #-NoId
            result1 = TEXTO.find('<ruc>')
            result2 = TEXTO.find('</ruc>')
            if (result1 != -1): 
                NoId=(TEXTO[result1+5:result2]) 
            else: 
                NoId=""
            #-AgenteRetencion
            result1 = TEXTO.find('<razonSocial>')
            result2 = TEXTO.find('</razonSocial>')
            if (result1 != -1): 
                AgenteRetencion=(TEXTO[result1+13:result2])
            else: 
                AgenteRetencion=""
            #-TipoComprobante
            result1 = TEXTO.find('<codDoc>')
            result2 = TEXTO.find('</codDoc>')
            if (result1 != -1): 
                TipoComprobante=(TEXTO[result1+8:result2])
            else: 
                TipoComprobante=""
            #-Establecimiento
            result1 = TEXTO.find('<estab>')
            result2 = TEXTO.find('</estab>')
            if (result1 != -1): 
                Establecimiento=(TEXTO[result1+7:result2])
            else: 
                Establecimiento=""
            #-PuntoEmision
            result1 = TEXTO.find('<ptoEmi>')
            result2 = TEXTO.find('</ptoEmi>')
            if (result1 != -1): 
                PuntoEmision=(TEXTO[result1+8:result2])
            else: 
                PuntoEmision=""
            #-Secuencial
            result1 = TEXTO.find('<secuencial>')
            result2 = TEXTO.find('</secuencial>')
            if (result1 != -1): 
                Secuencial=(TEXTO[result1+12:result2])
            else: 
                Secuencial=""
            #-ClaveAcceso
            result1 = TEXTO.find('<claveAcceso>')
            result2 = TEXTO.find('</claveAcceso>')
            if (result1 != -1): 
                ClaveAcceso=(TEXTO[result1+13:result2])
            else: 
                ClaveAcceso=""
            #-FechaEmision
            result1 = TEXTO.find('<fechaEmision>')
            result2 = TEXTO.find('</fechaEmision>')
            if (result1 != -1): 
                FechaEmision=(TEXTO[result1+14:result2])
            else: 
                FechaEmision=""

            #-TipoComprobSus
            result1 = TEXTO.find('<codDocSustento>')
            result2 = TEXTO.find('</codDocSustento>')
            if (result1 != -1): 
                TipoComprobSus=(TEXTO[result1+16:result2])
            else: 
                TipoComprobSus=""

            #-EstablecimientoSus-PtoEmisionSus-SecuencialSus
            result1 = TEXTO.find('<numDocSustento>')
            result2 = TEXTO.find('</numDocSustento>')
            numDocSustento=(TEXTO[result1+16:result2])
            if (result1 != -1): 
                EstablecimientoSus=(numDocSustento[0:3])
                PtoEmisionSus=(numDocSustento[3:6])
                SecuencialSus=(numDocSustento[6:15])
            else: 
                EstablecimientoSus1=""
                PtoEmisionSus=""
                SecuencialSus=""

            #-FechaEmisionSus
            result1 = TEXTO.find('<fechaEmisionDocSustento>')
            result2 = TEXTO.find('</fechaEmisionDocSustento>')
            if (result1 != -1): 
                FechaEmisionSus=(TEXTO[result1+25:result2])
            else: 
                FechaEmisionSus=""
            

            #-----------------------------------------------
            #IMPUESTOS
            #-----------------------------------------------
            #---------
            #--Codigo1
            trim1 = TEXTO.find('<codigo>')
            trim2 = trim1+450
            #trim2 = TEXTO.find('</impuesto>')
            #print(TEXTO.find('<codigo>'))
            if (trim1 != -1):
                contador+= 1
                Codigo1=(TEXTO[trim1+18:trim2])

                #-CodigoRetencion

                result1 = Codigo1.find('<codigoRetencion>')
                result2 = Codigo1.find('</codigoRetencion>')
                if (result1 != -1): 
                    CodigoRetencion1=(Codigo1[result1+17:result2])
                else: 
                    CodigoRetencion1=""

                #-BaseImponible

                result1 = Codigo1.find('<baseImponible>')
                result2 = Codigo1.find('</baseImponible>')
                if (result1 != -1): 
                    BaseImponible1=(Codigo1[result1+15:result2])
                else: 
                    BaseImponible1=""

                #-PorcentajeRetencion

                result1 = Codigo1.find('<porcentajeRetener>')
                result2 = Codigo1.find('</porcentajeRetener>')
                if (result1 != -1): 
                    PorcentajeRetencion1=(Codigo1[result1+19:result2])
                else: 
                    PorcentajeRetencion1=""

                #-ValorRetenido

                result1 = Codigo1.find('<valorRetenido>')
                result2 = Codigo1.find('</valorRetenido>')
                if (result1 != -1): 
                    ValorRetenido1=(Codigo1[result1+15:result2])
                else: 
                    ValorRetenido1=""
                    
            else: 
                    CodigoRetencion1="";BaseImponible1="";ValorRetenido1="";PorcentajeRetencion1=""

            #---------
            #--Codigo2
            trim1 = trim1+TEXTO[trim1+18:].find('<codigo>')
            trim2 = trim1+450
            #trim2 = trim1+TEXTO[trim1:].find('</impuesto>')
            if (TEXTO[trim1+8:].find('<codigo>') != -1):
                contador+=1
                Codigo2=(TEXTO[trim1+18:trim2])

                #-CodigoRetencion

                result1 = Codigo2.find('<codigoRetencion>')
                result2 = Codigo2.find('</codigoRetencion>')
                if (result1 != -1): 
                    CodigoRetencion2=(Codigo2[result1+17:result2])
                else: 
                    CodigoRetencion2=""

                #-BaseImponible

                result1 = Codigo2.find('<baseImponible>')
                result2 = Codigo2.find('</baseImponible>')
                if (result1 != -1): 
                    BaseImponible2=(Codigo2[result1+15:result2])
                else: 
                    BaseImponible2=""

                #-PorcentajeRetencion

                result1 = Codigo2.find('<porcentajeRetener>')
                result2 = Codigo2.find('</porcentajeRetener>')
                if (result1 != -1): 
                    PorcentajeRetencion2=(Codigo2[result1+19:result2])
                else: 
                    PorcentajeRetencion2=""

                #-ValorRetenido

                result1 = Codigo2.find('<valorRetenido>')
                result2 = Codigo2.find('</valorRetenido>')
                if (result1 != -1): 
                    ValorRetenido2=(Codigo2[result1+15:result2])
                else: 
                    ValorRetenido2=""

            else: 
                   CodigoRetencion2="";BaseImponible2="";ValorRetenido2="";PorcentajeRetencion2=""

            #---------
            #--Codigo3
            trim1 = 300+trim1+TEXTO[trim1+18:].find('<codigo>')
            trim2 = trim1+450
            #trim2 = trim2+TEXTO[trim1:].find('</impuesto>')
            if (TEXTO[trim1+8:].find('<codigo>')!= -1):
                contador+=1
                Codigo3=(TEXTO[trim1+18:trim2])

                #-CodigoRetencion

                result1 = Codigo3.find('<codigoRetencion>')
                result2 = Codigo3.find('</codigoRetencion>')
                if (result1 != -1): 
                    CodigoRetencion3=(Codigo3[result1+17:result2])
                else: 
                    CodigoRetencion3=""

                #-BaseImponible

                result1 = Codigo3.find('<baseImponible>')
                result2 = Codigo3.find('</baseImponible>')
                if (result1 != -1): 
                    BaseImponible3=(Codigo3[result1+15:result2])
                else: 
                    BaseImponible3=""

                #-PorcentajeRetencion

                result1 = Codigo3.find('<porcentajeRetener>')
                result2 = Codigo3.find('</porcentajeRetener>')
                if (result1 != -1): 
                    PorcentajeRetencion3=(Codigo3[result1+19:result2])
                else: 
                    PorcentajeRetencion3=""

                #-ValorRetenido

                result1 = Codigo3.find('<valorRetenido>')
                result2 = Codigo3.find('</valorRetenido>')
                
                if (result1 != -1): 
                    ValorRetenido3=(Codigo3[result1+15:result2])
                else: 
                    ValorRetenido3=""
                    
            else: 
                    CodigoRetencion3="";BaseImponible3="";ValorRetenido3="";PorcentajeRetencion3=""

            #---------
            #--Codigo4
            trim1 = 300+trim1+TEXTO[trim1+18:].find('<codigo>')
            trim2 = trim1+450
            if (TEXTO[trim1+8:].find('<codigo>') != -1):
                contador+=1
                Codigo4=(TEXTO[trim1+18:trim2])

                #-CodigoRetencion

                result1 = Codigo4.find('<codigoRetencion>')
                result2 = Codigo4.find('</codigoRetencion>')
                if (result1 != -1): 
                    CodigoRetencion4=(Codigo4[result1+17:result2])
                else: 
                    CodigoRetencion4=""

                #-BaseImponible

                result1 = Codigo4.find('<baseImponible>')
                result2 = Codigo4.find('</baseImponible>')
                if (result1 != -1): 
                    BaseImponible4=(Codigo4[result1+15:result2])
                else: 
                    BaseImponible4=""

                #-PorcentajeRetencion

                result1 = Codigo4.find('<porcentajeRetener>')
                result2 = Codigo4.find('</porcentajeRetener>')
                if (result1 != -1): 
                    PorcentajeRetencion4=(Codigo4[result1+19:result2])
                else: 
                    PorcentajeRetencion4=""

                #-ValorRetenido

                result1 = Codigo4.find('<valorRetenido>')
                result2 = Codigo4.find('</valorRetenido>')
                if (result1 != -1): 
                    ValorRetenido4=(Codigo4[result1+15:result2])
                else: 
                    ValorRetenido4=""

            else: 
                   CodigoRetencion4="";BaseImponible4="";ValorRetenido4="";PorcentajeRetencion4=""
            

            #Append final de elemento total extraido
            content = [ NoId,
                        AgenteRetencion,
                        TipoComprobante,
                        Establecimiento,
                        PuntoEmision,
                        Secuencial,
                        ClaveAcceso,
                        FechaEmision,
                        TipoComprobSus,
                        EstablecimientoSus,
                        PtoEmisionSus,
                        SecuencialSus,
                        FechaEmisionSus,
                        CodigoRetencion1,
                        BaseImponible1,
                        PorcentajeRetencion1,
                        ValorRetenido1,
                        CodigoRetencion2,
                        BaseImponible2,
                        PorcentajeRetencion2,
                        ValorRetenido2,
                        CodigoRetencion3,
                        BaseImponible3,
                        PorcentajeRetencion3,
                        ValorRetenido3,
                        CodigoRetencion4,
                        BaseImponible4,
                        PorcentajeRetencion4,
                        ValorRetenido4,
                        contador
                        ]

            elements.append(content)

#-------------------------------------------------------------------------------
        except:

            #print("------------------------------------------------")
            #print(f+"Failed")
            
            content = ["NULL", "NULL", "NULL", "NULL","NULL","NULL","NULL","NULL","NULL", "NULL",
                       "NULL", "NULL", "NULL", "NULL","NULL","NULL","NULL","NULL","NULL", "NULL",
                       "NULL", "NULL", "NULL", "NULL","NULL","NULL","NULL","NULL","NULL", "NULL"
                       ]

            elements.append(content)

            os.remove("testfile.txt")

    #//////////////////////////////////////////////////////
    #Escritura sobre Excel (Escribimos la lista elements en una hoja de excel)
    #//////////////////////////////////////////////////////

    name='LibroMayor_'+date+'.xlsx'
    
    if len(elements)>0:
        #Creación de Archivo Excel
        wb = xlsxwriter.Workbook(name)
          
        # add_sheet is used to create sheet. 
        sheet1 = wb.add_worksheet()
        # Start from the first cell. Rows and 
        # columns are zero indexed. 
        row = 0
        col = 0

        Headers = [ "No. Id.","Agente de Retención","Tipo Comprobante","Establecimiento","Punto Emisión","Secuencial","Clave de Acceso","Fecha Emisión",
                    "Tipo Comprob. Sus.","Establecimiento Sus.","Punto Emisión Sus.","Secuencial Sus.","Fecha Emisión Sus.",
                    "Cód. Ret. 1","Base Imponible 1","% Ret. 1","Valor Retenido 1",
                    "Cód. Ret. 2","Base Imponible 2","% Ret. 2","Valor Retenido 2",
                    "Cód. Ret. 3","Base Imponible 3","% Ret. 3","Valor Retenido 3",
                    "Cód. Ret. 4","Base Imponible 4","% Ret. 4","Valor Retenido 4",
                    "Contador","Status"]

        #Escritura de Headers
        for j in range(0,31):

           sheet1.write(0,j, Headers[j])
          
        # Iterate over the data and write it out row by row. 
        for x in (elements):

            for i in range(0, 30):

                sheet1.write(row+1, col+i, elements[row][col+i])
                
            row += 1
          
        wb.close()

    if os.path.isfile(name) == True:
    
        return "Archivo generado con exito"

    else:

        return "Error"

#readXML('C:/Users/Bryan/Desktop/Lector_de_XML/XML','18_9_2019')

if __name__ == "__main__":
    path=str(sys.argv[1])
    date=str(sys.argv[2])
    readXML(path,date)
    








