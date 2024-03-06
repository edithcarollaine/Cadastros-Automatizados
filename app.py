import openpyxl
import pyperclip
import pyautogui


workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

def copy_to_clipboard(text):
    pyperclip.copy(text)

for linha in sheet_produtos.iter_rows(min_row=2, max_row=2):
    nome_produto = linha[0].value
    copy_to_clipboard(nome_produto)
    pyautogui.click(24,179, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    descricao = linha[1].value
    copy_to_clipboard(descricao)
    pyautogui.click(23,248, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    categoria = linha[2].value
    copy_to_clipboard(categoria)
    pyautogui.click(21,312, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    codigo_produto = linha[3].value
    copy_to_clipboard(codigo_produto)
    pyautogui.click(27,384, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    peso = linha[4].value
    copy_to_clipboard(peso)
    pyautogui.click(32,453, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    dimensoes = linha[5].value
    copy_to_clipboard(dimensoes)
    pyautogui.click(11,524, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    ''' Caso possua um botão de próxima página '''
    # pyautogui.click(x,y, duration=1)
    # time.sleep(5)
    
    # preco = linha[6].value
    # quantidade_em_estoque = linha[7].value
    # data_de_validade = linha[8].value
    # cor = linha[9].value
    # tamanho = linha[10].value
    # material = linha[11].value
    # fabricante = linha[12].value
    # pais_origem = linha[13].value
    # observacoes = linha[14].value
    # codigo_de_barras = linha[15].value
    # localizacao_armazem = linha[16].value