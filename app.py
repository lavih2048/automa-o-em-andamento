import openpyxl
import pyperclip
import pyautogui
import time
#entrar na planilha

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
#copiar informação de um campo e colar no seu campo correspondente

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(427,210, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(462,297, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(429,470, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(430,531, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(421,591, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(417,652, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pyautogui.click(627,736, duration=1)
    time.sleep(1)
    
    # página 2
    
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(410,212, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(416,277, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(422,341, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(417,402, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    tamanho = linha[10].value
    pyperclip.copy(tamanho)
    pyautogui.click(412,465, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(411,528, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    #botão de próximo
    pyautogui.click(606,605, duration=1)
    time.sleep(1)
    
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(469,213, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(444,275, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(404,359, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(407,535, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(402,601, duration=1)
    pyautogui.hotkey('ctrl','v')
    
    
    #botão de próximo
    pyautogui.click(595,673, duration=1)
    pyautogui.click(847,161, duration=2)
    pyautogui.click(674,677, duration=2)
    time.sleep(2)
    
    
#repetir esses passos para outros campos até preencher campos daquela página
#clicar em próxima
#repetir os mesmos passos e ir para a próxima página (página 2)
#repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir
#clicar em ok, para finalizar processo
#cicar no ok mais uma vez na mensagem de confirmação de salvamento no banco de dados
#clicar em "adicionar mais um e repetir o processo até finalizar a planilha"