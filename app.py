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
    
    preco = linha[6].value
    copy_to_clipboard(preco)
    pyautogui.click(11, 524, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    quantidade_em_estoque = linha[7].value
    copy_to_clipboard(quantidade_em_estoque)
    pyautogui.click(11, 524, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    data_de_validade = linha[8].value
    copy_to_clipboard(data_de_validade)
    pyautogui.click(11, 524, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    cor = linha[9].value
    copy_to_clipboard(cor)
    pyautogui.click(11, 524, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    ''' 
    Caso tenha uma lista que tenha que selecionar algo
    ''' 
    tamanho = linha[10].value
    pyautogui.click(11,524, duration=1) 
    if tamanho == 'Pequeno':
        pyautogui.click(11, 524, duration=1)
    elif tamanho == 'Medio':
        pyautogui.click(11, 524, duration=1)
    else:
        pyautogui.click(11, 524, duration=1)
            
    material = linha[11].value
    copy_to_clipboard(material)
    pyautogui.click(11, 524, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    ''' Caso possua um botão de próxima página '''
    # pyautogui.click(x,y, duration=1)
    # time.sleep(5)    
    
    fabricante = linha[12].value
    copy_to_clipboard(fabricante)
    pyautogui.click(11, 524, duration=1)
    pyautogui.hotkey('ctrl', 'v')    
    
    pais_origem = linha[13].value
    copy_to_clipboard(pais_origem)
    pyautogui.click(11, 524, duration=1)
    pyautogui.hotkey('ctrl', 'v')    
    
    observacoes = linha[14].value
    copy_to_clipboard(observacoes)
    pyautogui.click(11, 524, duration=1)
    pyautogui.hotkey('ctrl', 'v')    
    
    codigo_de_barras = linha[15].value
    copy_to_clipboard(codigo_de_barras)
    pyautogui.click(11, 524, duration=1)
    pyautogui.hotkey('ctrl', 'v')    
    
    localizacao_armazem = linha[16].value
    copy_to_clipboard(localizacao_armazem)
    pyautogui.click(11, 524, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    ''' botão concluir cadastro'''
    # pyautogui.click(x,y, duration=1)
    '''botão confirmar inclusão caso exista'''
    # pyautogui.click(x,y, duration=1)
    '''botão confirmar inclusão 2 caso exista'''
    # pyautogui.click(x,y, duration=1)
    '''botão para iniciar novo cadastro'''
    # pyautogui.click(x,y, duration=1)
    # time.sleep(5)  
    
    