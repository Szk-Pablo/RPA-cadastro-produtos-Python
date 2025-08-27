import openpyxl
import pyperclip
import pyautogui
from time import sleep

# 1 - ENTRAR NA PLANILHA

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
pagina = workbook['Produtos']

# 2 - COPIAR INFORMAÇÃO DE UM CAMPO DA PLANILHA E COLAR NO CAMPO CORRESPONDENTE

for linha in pagina.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(596,348, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(622,437, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(624,566, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    cod_produto = linha[3].value
    pyperclip.copy(cod_produto)
    pyautogui.click(607,647, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(619,739, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    dimensao = linha[5].value
    pyperclip.copy(dimensao)
    pyautogui.click(631,819, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    #BOTAO PROXIMO
    pyautogui.click(126,879)
    sleep(5)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(594,369, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    qtd_estoque = linha[7].value
    pyperclip.copy(qtd_estoque)
    pyautogui.click(618,458, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    validade = linha[8].value
    pyperclip.copy(validade)
    pyautogui.click(612,542, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(603,628, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    tamanho = linha[10].value
    pyautogui.click(548,712, duration=1) 
    sleep(1)
    if tamanho == 'Pequeno':
        pyautogui.click(517,746, duration=1)
        sleep(1)
    elif tamanho == 'Médio':
        pyautogui.click(518,775, duration=1)
        sleep(1)
    else:
        pyautogui.click(522,803, duration=1)
        sleep(1)


    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(631,795, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    #BOTAO PROXIMO
    pyautogui.click(128,856, duration=1)
    sleep(5)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(612,387, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    origem = linha[13].value
    pyperclip.copy(origem)
    pyautogui.click(625,474, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(585,562, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    cod_barras = linha[15].value
    pyperclip.copy(cod_barras)
    pyautogui.click(630,695, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.click(609,778, duration=1)
    pyautogui.hotkey('ctrl','v')
    sleep(1)

    #BOTAO CONCLUIR
    pyautogui.click(126,838, duration=1)
    sleep(2)

    #BOTAO CONCLUIR 2
    pyautogui.click(824,190, duration=1)
    sleep(3)

    #BOTAO NOVO CADASTRO
    pyautogui.click(647,616, duration=1)
    sleep(3)
