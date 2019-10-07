import sys
import os

def hacerReporte(ID):

    f=open(ID+"_Recibidos.txt", "r+")

    TEXTO = f.read()

    TEXTO = TEXTO.replace('\t',';')
    TEXTO = TEXTO.replace('\n\n','\n')


    #f.save()
    f.close()

    file1 = open("sriReporte.csv","w+") 
    file1.write(TEXTO)
    file1.close()


if __name__ == "__main__":
    ID=str(sys.argv[1])
    hacerReporte(ID)



