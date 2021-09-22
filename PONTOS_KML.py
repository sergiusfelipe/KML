from tkinter import filedialog,ttk
import pandas as pds
from bs4 import BeautifulSoup
import csv
import tkinter as tk


def getKML():
    
    pl1 = filedialog.askopenfilename()
    data = pds.DataFrame()
    marcadores = pds.DataFrame()
    
    #Campos para armazenar os atributos dos pinos
    n_poste = []
    tipo = []
    t_r = []
    at = []
    mt = []
    bt = []
    ip = []
    chave = []
    trans = []
    atrr = []
    prop = []
    ctoe = []
    ceo = []
    qut_ctoe = []
    qut_ceo = []
    ocup = []
    casas = []
    coor = []
    marc = []
    ocup1 =[]
    ocup2 = []
    ocup3 = []
    ocup4 = []
    ocup5 = []
    ocup6 = []
    ocup7 = []
    bap = []
    alca = []
    laco = []
    sup = []
    rt = []
    
    with open(pl1, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f,'lxml')
        
        for m in soup.find_all('name'):
            marc.append(m)
        for coords in soup.find_all('coordinates'):
            coords = str(coords)
            coords.replace('<coordinates>','')
            coords.replace(',0</coordinates>','')
            coor.append(coords)
        '''for node in soup.find_all('description'):
            node = str(node)
            descri = pds.read_html(node)[0]
            print(descri)
            for i in range(0,23):
                if i == 0:
                    n_poste.append(descri.loc[i,1])
                elif i == 1:
                    tipo.append(descri.loc[i,1])
                elif i == 2:
                    t_r.append(descri.loc[i,1])
                elif i == 3:
                    at.append(descri.loc[i,1])
                elif i == 4:
                    mt.append(descri.loc[i,1])
                elif i == 5:
                    bt.append(descri.loc[i,1])
                elif i == 6:
                    ip.append(descri.loc[i,1])
                elif i == 7:
                    chave.append(descri.loc[i,1])
                elif i == 8:
                    trans.append(descri.loc[i,1])
                elif i == 9:
                    atrr.append(descri.loc[i,1])
                elif i == 10:
                    prop.append(descri.loc[i,1])
                elif i == 11:
                    ctoe.append(descri.loc[i,1])
                elif i == 12:
                    ceo.append(descri.loc[i,1])
                elif i == 13:
                    qut_ctoe.append(descri.loc[i,1])
                elif i == 14:
                    qut_ceo.append(descri.loc[i,1])
                elif i == 15:
                    ocup.append(descri.loc[i,1])
                elif i == 16:
                    ocup1.append(descri.loc[i,1])
                elif i == 17:
                    ocup2.append(descri.loc[i,1])
                elif i == 18:
                    ocup3.append(descri.loc[i,1])
                elif i == 19:
                    ocup4.append(descri.loc[i,1])
                elif i == 20:
                    ocup5.append(descri.loc[i,1])
                elif i == 21:
                    ocup6.append(descri.loc[i,1])
                elif i == 22:
                    casas.append(descri.loc[i,1])
                elif i == 23:
                    alca.append(descri.loc[i,1])
                elif i == 24:
                    laco.append(descri.loc[i,1])
                elif i == 25:
                    sup.append(descri.loc[i,1])
                elif i == 26:
                    rt.append(descri.loc[i,1])'''
        
    '''data['NumPoste'] = n_poste
    data['Tipo'] = tipo
    data['TamanhoResistencia'] = t_r
    data['AT'] = at
    data['MT'] = mt
    data['BT'] = bt
    data['IP'] = ip
    data['Chave'] = chave 
    data['Transformador'] = trans
    data['Aterramento'] = atrr
    data['Proprietario do poste'] = prop
    data['CTOE/TAR'] = ctoe
    data['CEO/ARMARIO'] = ceo
    data['QUANT CTOE'] = qut_ctoe
    data['QUANT/CEO'] = qut_ceo
    data['OCUPACOES'] = ocup
    data['OCUPANTE 1'] = ocup1
    data['OCUPANTE 2'] = ocup2
    data['OCUPANTE 3'] = ocup3
    data['OCUPANTE 4'] = ocup4
    data['OCUPANTE 5'] = ocup5
    data['OCUPANTE 6'] = ocup6'''
    #data['QTD BAP'] = bap
    #data['QTD ALCA'] = alca
    #data['QTD LACOS'] = laco
    #data['QTD SUPORTE'] = sup
    #data['RESERVA TECNICA'] = rt
    #data['OCUPANTE 7'] = ocup7
    #data['QUANT CASAS'] = casas
    data['COORDENADAS'] = coor
    marcadores['Marcadores'] = marc
    
    data['COORDENADAS'] = data['COORDENADAS'].astype(str)
    for i in range(0,len(data['COORDENADAS'])):
        a = data.loc[i,'COORDENADAS']
        a = a.replace('<coordinates>','')
        a = a.replace(',0</coordinates>','')
        a = a.strip()
        data.loc[i,'COORDENADAS'] = a
    
    data.to_excel('Coordenadas.xlsx')
    marcadores.to_excel('Marcadores.xlsx')
    concluido()

def end():
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

    browseButton_Excel_1 = tk.Button(text='EXTRAIR CSV', command=getKML, bg='white', fg='red', font=('helvetica', 12, 'bold'))
    canvas1.create_window(150, 150, window=browseButton_Excel_1)
    canvas1.create_text(125,295,fill="white",text="Desenvolvido por Sergio Tavora")

    root.mainloop()