import tkinter as tk
from tkinter import filedialog,ttk
import pandas as pd
import numpy as np
import simplekml
from polycircles import polycircles

def get1():
    pl1 = filedialog.askopenfilename()
    pl1_1 = pd.read_excel(pl1)
    pl1_2 = pd.DataFrame(pl1_1)
    
    return pl1_2

def GetKML():
    data = get1()
    kml = simplekml.Kml(name = 'NOME ARQUIVO', open = 1)
    HTML = pd.DataFrame()
    green = kml.newfolder(name='NOME PASTA')
    
    var3 = tk.DoubleVar()
    barra3 = ttk.Progressbar(root, variable = var3, maximum=len(data['Marcadores']),mode='determinate')
    barra3.pack()
    l1 = []
    l2 = []
    l3 = []
    
    j = 0
    k = 0
    l = 0
    
    j_1 = 0
    k_1 = 0
    l_1 = 0
    
    data['Marcadores'] = data['Marcadores'].map(str)
    
    for i in range(0, len(data['Marcadores'])):
        HTML = pd.DataFrame()
        HTML['CATEGORIA'] = ['CAIXA','LATITUDE','LONGITUDE']
        HTML['CONTEUDO'] = [data.loc[i,'Marcadores'],data.loc[i,'LATITUDE'],data.loc[i,'LONGITUDE']]
        circulo = polycircles.Polycircle(latitude = data.loc[i,'LATITUDE'], longitude = data.loc[i,'LONGITUDE'], radius = 70, number_of_vertices=36)
        KML = green.newpolygon(name=data.loc[i,'Marcadores'], outerboundaryis=circulo.to_kml())  # lon, lat, optional height
        #KML = green.newpoint(name=data.loc[i,'Marcadores'], coords = [(data.loc[i,'LONGITUDE'],data.loc[i,'LATITUDE'])] )
        KML.style.polystyle.color = '99FF0000' 
        table = HTML.to_html()
        KML.balloonstyle.text = table
        j = j + 1
        
        
        print('CARREGANDO: ',i*100/len(data['Marcadores']),'% . ',i+1,'pontos gerados')
        var3.set(i)
        root.update()
        
    total = l+1+k+1+j+1
    green.balloonstyle.text = 'Total de caixas: '+str(total)#+'. '+str((j+1)/total*100)+'% do total possuem status verde'
    
    kml.save("NOME ARQUIVO.kml")
    concluido()


def end ():
    return exit
    
def concluido():
    ok = tk.Toplevel(root)
    canvas2 = tk.Canvas(ok, width = 300, height = 300, bg = "#00ffff")
    botao = tk.Button(ok, text = 'Concluido', command = end, bg='white', fg='green', font=('helvetica', 12, 'bold'))
    canvas2.create_window(150, 150, window=botao)
    canvas2.pack()

if __name__ == "__main__":
    root= tk.Tk()

    canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'black')
    canvas1.pack()

    browseButton_Excel_1 = tk.Button(text='APLICAR RAIOS', command=GetKML, bg='white', fg='red', font=('helvetica', 12, 'bold'))
    canvas1.create_window(150, 150, window=browseButton_Excel_1)
    canvas1.create_text(125,295,fill="white",text="Desenvolvido por Sergio Tavora")
    
    root.mainloop()