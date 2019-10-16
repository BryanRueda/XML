#Importacion de librerias
import xml.etree.ElementTree as ET
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
        Obligacion=0
        DirEstablecimiento=0
        Establecimiento=0
        PuntoEmision=0
        Secuencial=0
        ClaveAcceso=0
        FechaEmision=0
        TipoComprobSus=0
        EstablecimientoSus=0
        PtoEmisionSus=0
        SecuencialSus=0
        FechaEmisionSus=0
        CodigoIMP1=0
        BaseImponible1=0
        Porcentaje1=0
        ValorRetenido1=0
        CodigoIMP2=0
        BaseImponible2=0
        Porcentaje2=0
        ValorRetenido2=0
        ORDEN=0

        Propina=0
        ImporteTotal=0
        FormaPago=0
        
        
        Codigo1=0
        Codigo2=0


        TEXTO=0
        contador=0
        trim1=0
        trim2=0
        result1=0
        result2=0

#-------------------------------------------------------------------------------
# Extracción de INFO
#-------------------------------------------------------------------------------

        try:

            #Lectura de archivo XML inicial
            mytree = ET.parse(f)
            myroot = mytree.getroot()

            #Extraccion de parametros uno a uno
                
            #-NoAutorizacion
            for x in myroot.findall('numeroAutorizacion'):
                NoAutorizacion=(x.text)

            #Extraer una cadena de texto con toda la informacion
            for x in myroot.findall('comprobante'):
                TEXTO=(x.text)
            
            #-NoId
            result1 = TEXTO.find('<ruc>')
            result2 = TEXTO.find('</ruc>')
            if (result1 != -1): 
                NoId=(TEXTO[result1+5:result2]) 
            else: 
                NoId=""

            #-ClaveAcceso
            result1 = TEXTO.find('<claveAcceso>')
            result2 = TEXTO.find('</claveAcceso>')
            if (result1 != -1): 
                ClaveAcceso=(TEXTO[result1+13:result2])
            else: 
                ClaveAcceso=""
                
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

            #-Direccion Establecimiento
            result1 = TEXTO.find('<dirMatriz>')
            result2 = TEXTO.find('</dirMatriz>')
            if (result1 != -1): 
                DirEstablecimiento=(TEXTO[result1+11:result2])
            else: 
                DirEstablecimiento=""


            
            #ORDEN
            result1 = TEXTO.find('<campoAdicional nombre="ORDEN">')
            ayuda=TEXTO[result1:result1+100]
            result2 = result1+ayuda.find('</campoAdicional>')
            
            if (result1 != -1): 
                ORDEN=(TEXTO[result1+32:result2])
            else: 
                ORDEN=""


            trim1 = TEXTO.find('<infoFactura>')
            trim2 = TEXTO.find('</infoFactura>')
            TEXTO = TEXTO[trim1:trim2]
            

            #-Contribuyente
            result1 = TEXTO.find('<contribuyenteEspecial>')
            result2 = TEXTO.find('</contribuyenteEspecial>')
            if (result1 != -1): 
                Contribuyente=(TEXTO[result1+23:result2])
            else: 
                Contribuyente=""

            #Obligacion
            result1 = TEXTO.find('<obligadoContabilidad>')
            result2 = TEXTO.find('</obligadoContabilidad>')
            if (result1 != -1): 
                Obligacion=(TEXTO[result1+22:result2])
            else: 
                Obligacion=""
            
            #-FechaEmision
            result1 = TEXTO.find('<fechaEmision>')
            result2 = TEXTO.find('</fechaEmision>')
            if (result1 != -1): 
                FechaEmision=(TEXTO[result1+14:result2])
            else: 
                FechaEmision=""

            #-Propina
            result1 = TEXTO.find('<propina>')
            result2 = TEXTO.find('</propina>')
            if (result1 != -1): 
                Propina=(TEXTO[result1+9:result2])
            else: 
                Propina=""
            
            #Importe Total
            result1 = TEXTO.find('<importeTotal>')
            result2 = TEXTO.find('</importeTotal>')
            if (result1 != -1): 
                ImporteTotal=(TEXTO[result1+14:result2])
            else: 
                ImporteTotal=""
                
            #Forma de Pago
            result1 = TEXTO.find('<formaPago>')
            result2 = TEXTO.find('</formaPago>')
            if (result1 != -1): 
                FormaPago=(TEXTO[result1+11:result2])
            else: 
                FormaPago=""


            TEXTO=TEXTO.replace(" ","")

            #-----------------------------------------------
            #IMPUESTOS
            #-----------------------------------------------
            #---------
            #--Codigo1
            trim1 = TEXTO.find('<codigoPorcentaje>2</codigoPorcentaje>')-56
            trim2 = trim1+400
            #trim2 = TEXTO.find('</impuesto>')            
            if (trim1 != -57):
                Codigo1=(TEXTO[trim1+18:trim2])

                #-CodigoIMP

                result1 = Codigo1.find('<codigo>')
                result2 = Codigo1.find('</codigo>')
                if (result1 != -1): 
                    CodigoIMP1=(Codigo1[result1+8:result2])
                else: 
                    CodigoIMP1=""

                #-BaseImponible

                result1 = Codigo1.find('<baseImponible>')
                result2 = Codigo1.find('</baseImponible>')
                if (result1 != -1): 
                    BaseImponible1=(Codigo1[result1+15:result2])
                else: 
                    BaseImponible1=""
                #print(BaseImponible1)

                #-Porcentaje

                result1 = Codigo1.find('<codigoPorcentaje>')
                result2 = Codigo1.find('</codigoPorcentaje>')
                if (result1 != -1): 
                    Porcentaje1=(Codigo1[result1+18:result2])
                else: 
                    Porcentaje1=""

                #-ValorRetenido

                result1 = Codigo1.find('<valor>')
                result2 = Codigo1.find('</valor>')
                if (result1 != -1): 
                    ValorRetenido1=(Codigo1[result1+7:result2])
                else: 
                    ValorRetenido1=""
                    
            else: 
                    CodigoIMP1="";BaseImponible1="";ValorRetenido1="";Porcentaje1=""
                    #print(NoAutorizacion, )

            #---------
            #--Codigo2
            trim1 = TEXTO.find('<codigoPorcentaje>0</codigoPorcentaje>')-56
            trim2 = trim1+400
            #trim2 = TEXTO.find('</impuesto>')
            
            if (trim1 != -57):
                Codigo2=(TEXTO[trim1+18:trim2])

                #-BaseImponible

                result1 = Codigo2.find('<baseImponible>')
                result2 = Codigo2.find('</baseImponible>')
                if (result1 != -1): 
                    BaseImponible2=(Codigo2[result1+15:result2])
                else: 
                    BaseImponible2=""


            else:
                BaseImponible2=""


            #Aplicar diccionario FORMAS DE PAGO
            formas=dict()
            formas={'01':"01-Sin utilización del sistema financiero",
                    '19':"19-Tarjeta de crédito",
                    '20':"20-Otros con utilización del sistema financiero"}

            FormaPago=formas.get(FormaPago)

            #Aplicacr diccionario 
            
            #Append final de elemento total extraido
            content = [ NoId,
                        TipoComprobante,
                        AgenteRetencion,
                        Contribuyente,
                        Obligacion,
                        DirEstablecimiento,
                        TipoComprobante,
                        Establecimiento,
                        PuntoEmision,
                        Secuencial,
                        NoAutorizacion,
                        ClaveAcceso,
                        FechaEmision,
                        "","","","","","","",
                        BaseImponible2,
                        BaseImponible1,
                        Porcentaje1,
                        ValorRetenido1,
                        Propina,
                        ImporteTotal,
                        FormaPago,
                        ORDEN
                        ]

            elements.append(content)

#-------------------------------------------------------------------------------
        except:

            #print("------------------------------------------------")
            #print(f+"Failed")
            
            content = ["NULL", "NULL", "NULL", "NULL","NULL","NULL","NULL","NULL","NULL", "NULL",
                       "NULL", "NULL", "NULL", "NULL","NULL","NULL","NULL","NULL","NULL", "NULL",
                       "NULL", "NULL", "NULL", "NULL","NULL","NULL","NULL"
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

        Headers = [ "No. Id.","Tipo Id.","Proveedor","Contrib. Especial","Oblig. Llevar Cont.","Direc. Establec.","Tipo Comprobante",
                    "Establec.","Pto. Emi.", "Secuencial", "Autorización", "Clave de Acceso", "Fecha Emisión",
                    "Tipo Comprob. Mod.","Establecimiento Mod.","Punto Emisión. Mod","Secuencial Mod.","Fecha Emi. Mod.",
                    "Base Exenta","Base No Objeto","Base Tarifa 0","Base Tarifa 12",
                    "Tarifa IVA","Monto IVA","Propina","Importe Total","Forma de Pago 1","ORDEN"]

        #Escritura de Headers
        for j in range(0,28):

           sheet1.write(0,j, Headers[j])
          
        # Iterate over the data and write it out row by row. 
        for x in (elements):

            for i in range(0, 28):

                sheet1.write(row+1, col+i, elements[row][col+i])
                
            row += 1
          
        wb.close()

    if os.path.isfile(name) == True:
    
        return "Archivo generado con exito"

    else:

        return "Error"

pat="C:/Users/bryan/Documents/DELOITTE/Lector_de_XML/CARLOS/XML/2019/10/"
day="20_9_2019"

readXML(pat,day)

##if __name__ == "__main__":
##    path=str(sys.argv[1])
##    date=str(sys.argv[2])
##    readXML(path,date)
##    








