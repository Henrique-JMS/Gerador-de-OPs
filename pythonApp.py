import FreeSimpleGUI as sg
from GerarOrdensDeProducao import gerarOrdemDeProducao

# Criação da interface de usuário, utilizando o pacote FreeSimpleGUI
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
    if event == sg.WIN_CLOSED or event == 'Exit':
        break      

    if event == '-BTNGERAR-':
        caminhoCarteira = values['-ARQUIVOSELECIONADO-']
        caminhoPasta = values['-PASTASELECIONADA-']
        gerarOrdemDeProducao(caminhoCarteira, caminhoPasta)
        

window.close()
