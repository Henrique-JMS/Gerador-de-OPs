import pandas as pd
import openpyxl
import numpy as np
import FreeSimpleGUI as sg
from pathlib import Path

# Carrega o arquivo que serve como modelo para a confecção das ordens de produção
caminhoOPModelo = Path(__file__).with_name('OP Modelo.xlsx')
opModelo = openpyxl.load_workbook(caminhoOPModelo)
opPlanilha = opModelo['Plan1']

def manipularCarteira(caminhoDaCarteira):
    # Cria um objeto DataFrame, do pacote Pandas, para a carteira.
    carteira = pd.read_excel(caminhoDaCarteira)

    # Remove primeira linha do índice, originalmente utilizada para a filtragem das colunas
    carteira = carteira.drop(index=0)

    # Remove linhas onde o saldo resultante é zero
    carteira = carteira[carteira['Saldo.1'] > 0]

    # Remove linhas onde o prazo não é definido e itens de tipo especial "k"
    carteira = carteira[~(carteira['Prazo'].str.contains('DF', na=False)) & ~(carteira['Prazo'].str.contains('S Con', na=False)) & ~(carteira['Tip'].str.contains('K', na=False))]

    # Altera o tipo de dado das células da coluna Prazo para numérico
    carteira['Prazo'] = carteira['Prazo'].astype(np.float64)

    # Altera o tipo de dado das células da coluna Data Embarque para texto
    carteira['Data Embarque'] = carteira['Data Embarque'].astype(str)

    # Filtra apenas as linhas onde o prazo é menor que doze semanas
    carteira = carteira[carteira.Prazo <= 12]

    # Remove linhas onde a criação da OP já foi registrada
    carteira = carteira[pd.isnull(carteira['OP'])]

    # Classifica os dados em ordem crescente por critério de prazo, código e ordem P
    carteira = carteira.sort_values(by=['Prazo', 'Código', 'OrdemP'])
    
    return carteira

def gerarOrdemDeProducao(caminhoDaCarteira, caminhoDaPasta):
    ct = 1
    carteiraOrganizada = manipularCarteira(caminhoDaCarteira)
    # A lista ordensPGeradas é utilizada como referência para atualizar a carteira.
    ordensPGeradas = []
    # Para cada item único, dado pelo código, da carteira que precisa de uma ordem de produção:
    for i in carteiraOrganizada.loc[:, 'Código'].drop_duplicates():
        
        # Planilha filtrada para todas as linhas do código específico
        carteiraFiltrada = carteiraOrganizada[carteiraOrganizada['Código'] == i]

        # Cria uma string com todos os pedidos e ordemP para o código
        listaDePedidos = carteiraFiltrada['Ped'].astype(str).tolist()
        listaPedidosRecortada = [w[:-2] for w in listaDePedidos]
        getped = ' | '.join(listaPedidosRecortada)
        listaDeOrdensP = carteiraFiltrada['OrdemP'].astype(str).tolist()
        listaOrdensPRecortada = [y[:-2] for y in listaDeOrdensP]
        getordemp = ' | '.join(listaOrdensPRecortada)

        ordensPGeradas.extend(listaOrdensPRecortada)

        # Utiliza slicing para remover a hora do padrão de datas e cria as strings para prazo, código, revisão e quantidade
        prazos = carteiraFiltrada['Data Embarque'].tolist()
        for i in range(len(prazos)):
            if not prazos[i] == 'Imediato':
                prazos[i] = prazos[i][0:10]
        getprazo = ' | '.join(prazos)
        getcodigo = carteiraFiltrada['Código'].iloc[0]
        getrevisao = carteiraFiltrada['Rev'].iloc[0]
        getquantidade = int(carteiraFiltrada['Saldo.1'].sum())


        # Preenche a ordem de produção modelo com as informações obtidas
        opPlanilha['O1'].value = ct
        opPlanilha['J2'].value = getcodigo
        opPlanilha['C6'].value = getprazo
        opPlanilha['C7'].value = getped
        opPlanilha['C8'].value = getordemp
        opPlanilha['Q2'].value = getrevisao
        opPlanilha['M6'].value = getquantidade

        # Informa o usuário que a OP está sendo gerada e salva o arquivo.
        print(f'Gerando OP: {getcodigo}...')
        opModelo.save(f'{caminhoDaPasta}/{ct} - {getcodigo}.xlsx')
        ct += 1

    # Gera uma carteira atualizada com os registros de criação das novas OPs
    carteira = openpyxl.load_workbook(caminhoDaCarteira)
    planilhaCarteira = carteira['Planilha1']
    linha = 1
    for celula in planilhaCarteira['A']:     
        if str(celula.value) in ordensPGeradas:
            celula.offset(row=0, column=11).value = 'OK'
        linha += linha       
    print("Salvando carteira atualizada...")
    carteira.save(f'{caminhoDaPasta}/Carteira atualizada.xlsx')
    print("Operação finalizada com sucesso.")


# Criação da interface de usuário, utilizando o pacote PySimpleGUI
sg.theme('GreenTan')
layout = [[sg.Text('Selecione a carteira de pedidos:')],      
          [sg.Input(key='-ARQUIVOSELECIONADO-', readonly=True), sg.FileBrowse()],
          [sg.Text('Selecione a pasta onde salvar os arquivos:')],    
          [sg.Input(key='-PASTASELECIONADA-', readonly=True), sg.FolderBrowse()],  
          [sg.Button('Gerar', key='-BTNGERAR-')],
          [sg.Output(size=(45,5))]]      

window = sg.Window('Gerar ordens de produção', layout)      

while True:           
    event, values = window.read() 
    #print(event, values)       
    if event == sg.WIN_CLOSED or event == 'Exit':
        break      

    if event == '-BTNGERAR-':
        caminhoCarteira = values['-ARQUIVOSELECIONADO-']
        caminhoPasta = values['-PASTASELECIONADA-']
        gerarOrdemDeProducao(caminhoCarteira, caminhoPasta)
        #print(event, values)

window.close()


