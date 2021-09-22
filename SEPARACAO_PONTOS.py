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
    
def mapa_reg():
    print("SELECIONAR PLANILHA")
    plan = get1()
    list_reg = plan['NomeCidade'].unique()
    for j in range(0,len(list_reg)):
        print(list_reg[j])
    for i in range(0,len(list_reg)):
        map_reg = plan[plan['NomeCidade'] == list_reg[i]]
        map_reg = map_reg.reset_index(drop=True)
        print(list_reg[i])
        #try:
        GetKML(str(list_reg[i]),map_reg)
        #except:
            #print("ERRO")
    concluido()
    

def GetKML():
    data = get1()
    data = data.sort_values('bairro')
    data = data.reset_index()
    nome_completo = 'NOME ARQUIVO'
    kml = simplekml.Kml(name = nome_completo, open = 1)
    HTML = pd.DataFrame()
    green = kml.newfolder(name='NOME PASTA')
    green_1 = green.newfolder(name = str(data.loc[1,'bairro']),open = 1)
    
    var3 = tk.DoubleVar()
    barra3 = ttk.Progressbar(root, variable = var3, maximum=len(data['COMPLETO']),mode='determinate')
    barra3.pack()
    
    l1 = []
    l2 = []
    l3 = []
    l4 = []
    l5 = []
    
    j = 0
    k = 0
    l = 0
    m = 0
    n = 0
    
    j_1 = 0
    k_1 = 0
    l_1 = 0
    m_1 = 0
    n_1 = 0
    
    for i in range(0, len(data['COMPLETO'])):
        HTML = pd.DataFrame()
        HTML['CATEGORIA'] = ['ID','NOME','LATITUDE','LONGITUDE']
        HTML['CONTEUDO'] = [data.loc[i,'COMPLETO'],data.loc[i,'nome_razaosocial'],data.loc[i,'LAT2'],data.loc[i,'LON2']]
        #circulo = polycircles.Polycircle(latitude = data.loc[i,'LATITUDE'], longitude = data.loc[i,'LONGITUDE'], radius = 70, number_of_vertices=36)
        l1.append(str(data.loc[i,'bairro']))
        if l1[j] != l1[j-1]:
            green_1.balloonstyle.text = 'Total de Clientes: '+str(j-j_1)+'.'
            green_1 = green.newfolder(name = data.loc[i,'bairro'],open = 1)
            j_1 = j
        
        KML = green_1.newpoint(name=data.loc[i,'nome_razaosocial'], coords = [(data.loc[i,'LON2'],data.loc[i,'LAT2'])] )  # lon, lat, optional height
        KML.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
        KML.style.labelstyle.scale = 0.1
        #KML.style.iconstyle.color = 'ff00ff00' 
        table = HTML.to_html()
        KML.balloonstyle.text = table
        KML.visibility = 0
        j = j + 1
        
        
        print('CARREGANDO: ',i*100/len(data['COMPLETO']),'% . ',i+1,'pontos gerados')
        var3.set(i)
        root.update()
        
    
    nome1 = nome_completo + ".kml"
    kml.save(nome1)


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

    browseButton_Excel_1 = tk.Button(text='PLOTAR POSTES', command=GetKML, bg='white', fg='red', font=('helvetica', 12, 'bold'))
    canvas1.create_window(150, 150, window=browseButton_Excel_1)
    canvas1.create_text(125,295,fill="white",text="Desenvolvido por Sergio Tavora")

    root.mainloop()