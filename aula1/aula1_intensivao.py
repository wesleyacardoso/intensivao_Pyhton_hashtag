import pyautogui
import time
import pyperclip
import pandas as pd

pyautogui.PAUSE = 3

pyautogui.press("winleft")
pyautogui.write("chrome")
pyautogui.press("enter")
time.sleep(3)
#pyautogui.alert("Vai começar, aperte OK e não mexa em nada.")
#pyautogui.hotkey('ctrl', 't') # nova guia do chrome
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga") # endereço aonde arquivo está
pyautogui.hotkey('ctrl', 'v') # usado para poder copiar posiveis caracteres especiais
# usando esses dois (pyperclip.copy=faz a copia do jeito que foi digitado,pyautogui.hotkey=executa tecla atalho)
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(x=340, y=300, clicks=2) # posição da pasta exportar

time.sleep(2)
pyautogui.click(x=376, y=372) # posição planilha
pyautogui.click(x=1163, y=191) # posição tres pontinhos
pyautogui.click(x=932, y=626) # poisção download
time.sleep(5)
# aqui abaixo se faz a leitura da planilha e os calculos usando a coluna VALOR FINAL e QUANTIDADE
tabela = pd.read_excel(r"C:\Users\wesle\Downloads\Vendas - Dez.xlsx")
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
# aqui abre o email
pyautogui.hotkey('ctrl', 't')
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")
pyautogui.hotkey('ctrl', 'v')
pyautogui.press("enter")
time.sleep(5)

pyautogui.click(x=78, y=201) # aqui é a posiçãp ESCREVER do email e quando abre
# o curso já vai direto para campo destinatario
pyautogui.write("wesleyacardoso@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
# aqui é o texto do corpo do email
texto = f"""
Prezados, bom dia!

O faturamento de ontem foi de : R$ {faturamento:,.2f}
A quantidade de produtos foi de : {quantidade:,}

Abs
Wesley Cardoso
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")  # envia o email























