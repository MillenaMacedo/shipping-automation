#!/usr/bin/env python
# coding: utf-8

# In[25]:


import pyperclip
import time
import pyautogui



#entrar no drive
pyautogui.PAUSE=1
pyautogui.hotkey("ctrl", "t")
pyperclip.copy ("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(4)
pyautogui.click(x=377, y=319, clicks=2)

# Baixar o arquivo
time.sleep(4)
pyautogui.click(x=404, y=478)
pyautogui.click(x=1167, y=203)
pyautogui.click(x=973, y=605)
time.sleep(3)
pyautogui.click(x=67, y=215)
time.sleep(5)
pyautogui.click(x=332, y=295)







#ler arquivo com pandas
import pandas as pd
df = pd.read_excel(r"C:\Users\mille\Downloads\Vendas - Dez.xlsx")
display(df)

faturamento = df["Valor Final"].sum()
qtdd_produtos = df["Quantidade"].sum()



#enviar o e-mail
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(4)
pyautogui.click(x=90, y=215)
pyperclip.copy("pythonimpressionador+diretoria@gmail.com")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
pyautogui.press("tab")
Assunto = "Relatório de vendas"
pyperclip.copy(Assunto)
pyautogui.hotkey("ctrl","v")
pyautogui.press("tab")
texto = f"""
Boa noite, prezados
Os resultados das vendas de hoje foram
Faturamento = {faturamento: ,.2f}
Quantidade de produtos vendidos = {qtdd_produtos:,}

Att, Millena."""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")
pyautogui.hotkey("ctrl","enter")

