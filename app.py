import pyautogui
import openpyxl
import pyperclip

# site para teste: https://cadastro-produtos-devaprender.netlify.app/index.html

# abrir planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
# carregar tabela
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    # copia e cola nome do produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(859, 226, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # copia e cola a descricao
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(844, 311, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # copia e cola a categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(844, 414, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # copia e cola o codigo do produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(844, 467, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # copia e cola o peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(824, 537, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # copia e cola as dimensoes
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(825, 611, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # clica no botão proximo
    pyautogui.click(833, 658, duration=1)

    # copia e cola o preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(819, 246, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # copia e cola a quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(819, 316, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # copia e cola a validade
    validade = linha[8].value
    pyperclip.copy(validade)
    pyautogui.click(819, 385, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # copia e cola a cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(819, 458, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # clica no tamanho, compara o valor copiado e seleciona o valor
    tamanho = linha[10].value
    pyperclip.copy(tamanho)
    pyautogui.click(819, 527, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(819, 551, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(819, 567, duration=1)
    else:
        pyautogui.click(819, 584, duration=1)

    # copia e cola o material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(819, 593, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # clica no próximo
    pyautogui.click(819, 648, duration=1)

    # copia e cola o fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(819, 266, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # copia e cola o país de origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(819, 335, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # copia e cola as observaões
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(819, 401, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # copia e cola o código de barras
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(819, 507, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # copia e cola a localização do armazem
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(819, 576, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # clica em concluir
    pyautogui.click(819, 622, duration=1)

    # clica em OK
    pyautogui.click(1192, 187, duration=1)

    # clica em Adicionar outro
    pyautogui.click(991, 439, duration=1)
