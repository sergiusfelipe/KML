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
    kml = simplekml.Kml(name = 'EXPANSAO/AMPLIACAO', open = 1)
    HTML = pd.DataFrame()
    green = kml.newfolder(name='EXPANSAO')
    yellow = kml.newfolder(name='AMPLIACAO')
    
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
    
    for i in range(0, len(data['Marcadores'])):
        HTML = pd.DataFrame()
        HTML['CATEGORIA'] = ['Marcadores','LATITUDE','LONGITUDE','TIPO']
        HTML['CONTEUDO'] = [data.loc[i,'Marcadores'],data.loc[i,'LATITUDE'],data.loc[i,'LONGITUDE'],data.loc[i,'COR']]
        if data.loc[i,'COR'] == 'EXPANSAO':
            KML = green.newpoint(name=data.loc[i,'Marcadores'], coords = [(data.loc[i,'LONGITUDE'],data.loc[i,'LATITUDE'])] )
            KML.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/polygon.png'
            KML.style.iconstyle.color = 'ff00ff00' 
            table = HTML.to_html()
            KML.balloonstyle.text = table
            j = j + 1
        elif data.loc[i,'COR'] == 'AMPLIACAO':
            KML = yellow.newpoint(name=data.loc[i,'Marcadores'], coords = [(data.loc[i,'LONGITUDE'],data.loc[i,'LATITUDE'])] )
            KML.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/polygon.png'
            KML.style.iconstyle.color = 'ff00ffff'
            table = HTML.to_html()
            KML.balloonstyle.text = table
            k = k + 1
        
        
        print('CARREGANDO: ',i*100/len(data['Marcadores']),'% . ',i+1,'pontos gerados')
        var3.set(i)
        root.update()
        
    kml.save("EXPANSAO_AMPLIACAO - BBB.kml")


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

    canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'red')
    canvas1.pack()

    browseButton_Excel_1 = tk.Button(text='EXPANSAO/AMPLIACAO', command=GetKML, bg='white', fg='red', font=('helvetica', 12, 'bold'))
    canvas1.create_window(150, 150, window=browseButton_Excel_1)
    canvas1.create_text(125,295,fill="white",text="Desenvolvido por Sergio Tavora")

    root.mainloop()