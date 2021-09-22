import tkinter as tk
from tkinter import filedialog,ttk
import pandas as pd
import numpy as np
import simplekml
from polycircles import polycircles

def mapa_geral():
    
    plan = getExcel()
    #plan = get1()
    GetKML("NOME AQUIVO",plan)
    
    concluido()
    
    
def mapa_reg():
    plan = getExcel()
    #plan = get1()
    print("SELECIONAR PLANILHA DE REGIONAIS")
    regionais = get1()
    list_reg = regionais['REGIONAL'].unique()
    for j in range(0,len(list_reg)):
        print(list_reg[j])
    cruzamento = pd.merge(plan,regionais,on='CIDADE',how='left')
    cruzamento.to_excel('PLANILHA_MAPA_DE_CALOR.xlsx')
    for i in range(0,len(list_reg)):
        map_reg = cruzamento[cruzamento['REGIONAL'] == list_reg[i]]
        map_reg = map_reg.reset_index(drop=True)
        print(list_reg[i])
        try:
            GetKML(str(list_reg[i]),map_reg)
        except:
            print("ERRO")
    concluido()

def get1():
    pl1 = filedialog.askopenfilename()
    pl1_1 = pd.read_excel(pl1)
    pl1_2 = pd.DataFrame(pl1_1)
    
    return pl1_2

def GetKML(nome,data):
    print(data)
    nome1 = "MAPA DE CALOR - " + str(nome)
    kml = simplekml.Kml(name = nome1, open = 1)
    HTML = pd.DataFrame()
    green = kml.newfolder(name='VERDE - 0% a 70%')
    #green.description = 'VERDE - 0% a 70%'
    red = kml.newfolder(name='VERMELHO - 90,1% a 100%')
    #red.description = 'VERMELHO - 90,1% a 100%'
    yellow = kml.newfolder(name='AMARELO - 70,1% a 90%')
    #yellow.description = 'AMARELO - 70,1% a 90%'
    purple = kml.newfolder(name='REDE BLOQUEADA')
    
    blue = kml.newfolder(name='AMPLIADO')
    
    var3 = tk.DoubleVar()
    barra3 = ttk.Progressbar(root, variable = var3, maximum=len(data['NOME_LOCAL']),mode='determinate')
    barra3.pack()
    green_1 = green.newfolder(name = str(data.loc[1,'CIDADE']))
    red_1 = red.newfolder(name = str(data.loc[1,'CIDADE']))
    yellow_1 = yellow.newfolder(name = str(data.loc[1,'CIDADE']))
    purple_1 = purple.newfolder(name = str(data.loc[1,'CIDADE']))
    blue_1 = blue.newfolder(name = str(data.loc[1,'CIDADE']))
    
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
    
    for i in range(0, len(data['NOME_LOCAL'])):
        HTML = pd.DataFrame()
        HTML['CATEGORIA'] = ['ID_CAIXA','NOME_LOCAL','DESCRICAO','TICKET_PROJETO','PROJETO','LATITUDE','LONGITUDE','CAPACIDADE','OCUPACAO','ESTADO','CIDADE','NOME_SPLITTER','SUBCATEGORIA_SPLITTER','COR']
        HTML['CONTEUDO'] = [data.loc[i,'ID_CAIXA'],data.loc[i,'NOME_LOCAL'],data.loc[i,'DESCRICAO'],data.loc[i,'TICKET_PROJETO'],data.loc[i,'PROJETO'],data.loc[i,'LATITUDE'],data.loc[i,'LONGITUDE'],data.loc[i,'CAPACIDADE'],data.loc[i,'OCUP'],data.loc[i,'ESTADO'],data.loc[i,'CIDADE'],data.loc[i,'NOME_SPLITTER'],data.loc[i,'SUBCATEGORIA_SPLITTER'],data.loc[i,'COR']]
        circulo = polycircles.Polycircle(latitude = data.loc[i,'LATITUDE'], longitude = data.loc[i,'LONGITUDE'], radius = 70, number_of_vertices=36)
        if data.loc[i,'COR'] == 'VERDE - 0% a 70%':
            l1.append(str(data.loc[i,'CIDADE']))
            if l1[j] != l1[j-1]:
                green_1.balloonstyle.text = 'Total de caixas: '+str(j-j_1)+' verdes.'
                green_1 = green.newfolder(name = data.loc[i,'CIDADE'],open = 1)
                j_1 = j
            KML = green_1.newpolygon(name=data.loc[i,'NOME_LOCAL'], outerboundaryis=circulo.to_kml())  # lon, lat, optional height
            KML.style.polystyle.color = '9914F000' 
            table = HTML.to_html()
            KML.balloonstyle.text = table
            j = j + 1
        elif data.loc[i,'COR'] == 'AMARELO - 70,1% a 90%':
            l2.append(str(data.loc[i,'CIDADE']))
            if l2[k] != l2[k-1]:
                yellow_1.balloonstyle.text = 'Total de caixas: '+str(k-k_1)+' amarelas.'
                yellow_1 = yellow.newfolder(name = data.loc[i,'CIDADE'],open = 1)
                k_1 = k
            KML = yellow_1.newpolygon(name=data.loc[i,'NOME_LOCAL'], outerboundaryis=circulo.to_kml())  # lon, lat, optional height
            KML.style.polystyle.color = '9914F0FF'
            table = HTML.to_html()
            KML.balloonstyle.text = table
            k = k + 1
        elif data.loc[i,'COR'] == 'VERMELHO - 90,1% a 100%':
            l3.append(str(data.loc[i,'CIDADE']))
            if l3[l] != l3[l-1]:
                red_1.balloonstyle.text = 'Total de caixas: '+str(l-l_1)+' vermelhas.'
                red_1 = red.newfolder(name = data.loc[i,'CIDADE'],open = 1)
                l_1 = l
            KML = red_1.newpolygon(name=data.loc[i,'NOME_LOCAL'], outerboundaryis=circulo.to_kml())  # lon, lat, optional height
            KML.style.polystyle.color = '990000ff' #990000ff
            table = HTML.to_html()
            KML.balloonstyle.text = table
            l = l + 1
        elif data.loc[i,'COR'] == 'REDE BLOQUEADA':
            l4.append(str(data.loc[i,'CIDADE']))
            if l4[m] != l4[m-1]:
                purple_1.balloonstyle.text = 'Total de caixas: '+str(m-m_1)+' bloqueadas.'
                purple_1 = purple.newfolder(name = data.loc[i,'CIDADE'],open = 1)
                m_1 = m
            KML = purple_1.newpolygon(name=data.loc[i,'NOME_LOCAL'], outerboundaryis=circulo.to_kml())  # lon, lat, optional height
            KML.style.polystyle.color = '99800080' #990000ff
            table = HTML.to_html()
            KML.balloonstyle.text = table
            m = m + 1
        elif data.loc[i,'COR'] == 'AMPLIADO':
            l5.append(str(data.loc[i,'CIDADE']))
            if l5[n] != l5[n-1]:
                blue_1.balloonstyle.text = 'Total de caixas: '+str(n-n_1)+' ampliadas.'
                blue_1 = blue.newfolder(name = data.loc[i,'CIDADE'],open = 1)
                n_1 = n
            KML = blue_1.newpolygon(name=data.loc[i,'NOME_LOCAL'], outerboundaryis=circulo.to_kml())  # lon, lat, optional height
            KML.style.polystyle.color = '99FF0000' #990000ff
            table = HTML.to_html()
            KML.balloonstyle.text = table
            n = n + 1
        
        print('CARREGANDO: ',i*100/len(data['NOME_LOCAL']),'% . ',i+1,'pontos gerados')
        var3.set(i)
        root.update()
        
    
    green_1.balloonstyle.text = 'Total de caixas: '+str(j-j_1)+' verdes.'
    yellow_1.balloonstyle.text = 'Total de caixas: '+str(k-k_1)+' amarelas.'
    red_1.balloonstyle.text = 'Total de caixas: '+str(l-l_1)+' vermelhas.'
    purple_1.balloonstyle.text = 'Total de caixas: '+str(m-m_1)+' bloqueados.'
    blue_1.balloonstyle.text = 'Total de caixas: '+str(n-n_1)+' ampliados.'
    
    total = l+1+k+1+j+1+m+1+n+1
    
    green.balloonstyle.text = 'Total de caixas: '+str(total)+'. '+str((j+1)/total*100)+'% do total possuem status verde'
    red.balloonstyle.text = 'Total de caixas: '+str(total)+'. '+str((l+1)/total*100)+'% do total possuem status vermelho'
    yellow.balloonstyle.text = 'Total de caixas: '+str(total)+'. '+str((k+1)/total*100)+'% do total possuem status amarelo'
    purple.balloonstyle.text = 'Total de caixas: '+str(total)+'. '+str((m+1)/total*100)+'% do total sao bloaqueados'
    blue.balloonstyle.text = 'Total de caixas: '+str(total)+'. '+str((n+1)/total*100)+'% do total sao ampliados'
    
    nome1 = nome + ".kml"
    kml.save(nome1)

def end ():
    return exit
    
def concluido():
    ok = tk.Toplevel(root)
    canvas2 = tk.Canvas(ok, width = 300, height = 300, bg = "#00ffff")
    botao = tk.Button(ok, text = 'Concluido', command = end, bg='white', fg='green', font=('helvetica', 12, 'bold'))
    canvas2.create_window(150, 150, window=botao)
    canvas2.pack()

def getExcel ():
    global df
    
    import_file_path = filedialog.askopenfilename()
    df = pd.read_excel (import_file_path)
    #df1 = pd.DataFrame(df, columns=['NOME_LOCAL','STATUS_LOCAL','LATITUDE','LONGITUDE','CAPACIDADE','PONTO_STATUS','ID_CAIXA','TICKET_PROJETO','DESCRICAO','SUBCATEGORIA_SPLITTER','NOME_SPLITTER','PROJETO','ESTADO','CIDADE'])
    
    df2 = df.groupby('NOME_LOCAL').agg(lambda x: list(x)).reset_index()
    
    ocup = []
    cap = []
    lat = []
    lon = []
    cor  = []
    ID = []
    TICKET = []
    DES = []
    SP = []
    SUB = []
    PRO = []
    ES = []
    CID = []
    local = []
    STATUS = []
    COND = []
    AMP = []
    
    
    var1 = tk.DoubleVar()
    barra1 = ttk.Progressbar(root, variable = var1, maximum=len(df2['PONTO_STATUS']),mode='determinate')
    barra1.pack()
    
    for i in range(0,len(df2['PONTO_STATUS'])):
      
        elemento2 = df2.loc[i,'LATITUDE']
        elemento3 = df2.loc[i,'LONGITUDE']
        elemento4 = df2.loc[i,'ID_CAIXA']
        elemento5 = df2.loc[i,'TICKET_PROJETO']
        elemento6 = df2.loc[i,'DESCRICAO']
        elemento7 = df2.loc[i,'NOME_SPLITTER']
        elemento8 = df2.loc[i,'SUBCATEGORIA_SPLITTER']
        elemento9 = df2.loc[i,'PROJETO']
        elemento10 = df2.loc[i,'ESTADO']
        elemento11 = df2.loc[i,'CIDADE']
        elemento12 = df2.loc[i,'STATUS_LOCAL']
        elemento13 = df2.loc[i,'CONDOMINIO']
        elemento14 = df2.loc[i,'CAIXA_AMPLIADA']
        t = df2.loc[i,'PONTO_STATUS'].count('Ocupado') + df2.loc[i,'PONTO_STATUS'].count('Aguardando Retirada') + df2.loc[i,'PONTO_STATUS'].count('Reservado')
        ocup.append(t)
        cap.append(len(df2.loc[i,'CAPACIDADE']))
        lat.append(max(elemento2))
        lon.append(max(elemento3))
        ID.append(max(elemento4))
        TICKET.append(max(elemento5))
        DES.append(max(elemento6))
        SP.append(max(elemento7))
        SUB.append(max(elemento8))
        PRO.append(max(elemento9))
        ES.append(max(elemento10))
        CID.append(max(elemento11))
        local.append(max(elemento12))
        COND.append(max(elemento13))
        AMP.append(max(elemento14))
        print('CARREGANDO 1 de 3: ',i*100/len(df2['PONTO_STATUS']),'%')
        var1.set(i)
        root.update()
    
    
    df2['OCUP'] = ocup
    df2['CAPACIDADE'] = cap
    df2['LATITUDE'] = lat
    df2['LONGITUDE'] = lon
    df2['ID_CAIXA'] = ID
    df2['CIDADE'] = CID
    df2['TICKET_PROJETO'] = TICKET
    df2['DESCRICAO'] = DES
    df2['NOME_SPLITTER'] = SP
    df2['SUBCATEGORIA_SPLITTER'] = SUB
    df2['PROJETO'] = PRO
    df2['CIDADE'] = CID
    df2['ESTADO'] = ES
    df2['STATUS_LOCAL'] = local
    df2['CONDOMINIO'] = COND
    df2['CAIXA_AMPLIADA'] = AMP
    
    ocupacao = []
    df2['LIVRE'] = df2['CAPACIDADE'] - df2['OCUP']
    df2['OCUPACAO(%)'] = df2['OCUP'] / df2['CAPACIDADE']
    
    var2 = tk.DoubleVar()
    barra2 = ttk.Progressbar(root, variable = var2, maximum=len(df2['PONTO_STATUS']),mode='determinate')
    barra2.pack()
    
    for j in range(0,len(df2['NOME_LOCAL'])):
        if df2.loc[j,'CAIXA_AMPLIADA'] == 'SIM':
            cor.append('AMPLIADO')
        elif df2.loc[j,'STATUS_LOCAL'] == 'Bloqueada' or df2.loc[j,'STATUS_LOCAL'] == 'Auditoria':
            cor.append('REDE BLOQUEADA')
        elif df2.loc[j,'OCUPACAO(%)'] <= 0.7 and df2.loc[j,'STATUS_LOCAL'] != 'Bloqueada':
            cor.append('VERDE - 0% a 70%')
        elif 0.7 < df2.loc[j,'OCUPACAO(%)'] <= 0.9 and df2.loc[j,'STATUS_LOCAL'] != 'Bloqueada':
            cor.append('AMARELO - 70,1% a 90%')
        elif 0.9 < df2.loc[j,'OCUPACAO(%)'] <= 1 and df2.loc[j,'STATUS_LOCAL'] != 'Bloqueada':
            cor.append('VERMELHO - 90,1% a 100%')
        elif df2.loc[j,'STATUS_LOCAL'] == 'Bloqueada':
            cor.append('REDE BLOQUEADA')
        else:
            df2.loc[j,'OCUPACAO(%)'] = 1
            cor.append('VERMELHO - 90,1% a 100%')
        
        print('CARREGANDO 2 de 3: ',j*100/len(df2['NOME_LOCAL']),'%')
        var2.set(j)
        root.update()
    
    df2['COR'] = cor
    df2['NOME_LOCAL'] = df2['NOME_LOCAL'].astype(str)
    df2['NOME_LOCAL'] = df2['NOME_LOCAL'].str.strip()
    df2.to_excel('PLANILHA_MAPA_DE_CALOR.xlsx')
    
    return df2
    
if __name__ == "__main__": 

    root= tk.Tk()

    canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'black')
    canvas1.pack()
    
    browseButton_Excel_1 = tk.Button(text='GERAL', command=mapa_geral, bg='white', fg='red', font=('helvetica', 12, 'bold'))
    browseButton_Excel_2 = tk.Button(text='REGIONAIS', command=mapa_reg, bg='white', fg='red', font=('helvetica', 12, 'bold'))
    canvas1.create_window(150, 100, window=browseButton_Excel_1)
    canvas1.create_window(150, 200, window=browseButton_Excel_2)
    canvas1.create_text(125,295,fill="white",text="Desenvolvido por Sergio Tavora")

    root.mainloop()