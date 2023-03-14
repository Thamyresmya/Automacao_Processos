import pyautogui
import pyperclip
import time
import pandas
pyautogui.PAUSE = 1

# Passo 1: Navegar até o local do relatorio
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://drive.google.com/drive/folders/1x9noGgNNFu7UQw4wbmGDlcr9ejpKIasj")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")

time.sleep(5)

# Passo 2: Navegar até o local do relatorio (entrar na pasta Exportar)
pyautogui.click(x=361, y=291, clicks=2)
time.sleep(2)

# Passo 3: Exportar o relatorio
pyautogui.click(x=351, y=286)
pyautogui.click(x=1719, y=191)
pyautogui.click(x=1572, y=532)
time.sleep(2)

# Passo 4: Calcular os indicadores (faturamento e quantidade de produtos)
tabela = pandas.read_excel(r"C:\Users\thamy\Downloads\Vendas - Dez.xlsx")
print(tabela)

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

print(faturamento)
print(quantidade)

# Passo 5:Enviar um e-mail para a diretoria
# entrar no email
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=om#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(2)

# nova msg
pyautogui.click(x=96, y=185)

# destinatario
pyautogui.write("miathamyres@gmail.com")
pyautogui.press("tab")

#assunto
pyautogui.press("tab")
pyperclip.copy("Relatório Faturamento do dia")
pyautogui.hotkey("ctrl","v")

# corpo de e-mail
pyautogui.press("tab")
texto = f""" 
    Prezados,
    
    O faturamento de ontem foi de: R$ {faturamento:,.2f}
    A quantidade de produtos foi de: {quantidade:,}
    
    Abs
    Thamyres Cavalcante    
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")

# enviar email
pyautogui.hotkey("ctrl","enter")


# Pegar localização
#print(pyautogui.position())





