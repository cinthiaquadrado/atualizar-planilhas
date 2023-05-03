import os
import time
import pyautogui
from openpyxl import load_workbook

folders = [r"C:\Users\cinth\Desktop\Downloads\pasta1", r"C:\Users\cinth\Desktop\Downloads\pasta2"]

file_names = ["teste1.xlsx", "teste1 - Copia - Copia.xlsx"]

for i, file_name in enumerate(file_names):
  os.chdir(folders[i])

# Use o PyAutoGUI para abrir o arquivo de Excel
  pyautogui.hotkey('ctrl', 'o')  # atalho para abrir o menu de abrir arquivo
  pyautogui.typewrite(file_name)  # digita o nome do arquivo
  pyautogui.press('enter')  # pressiona a tecla enter para abrir o arquivo
        
tab_order = [
    ["Planilha1"],
    ["Planilha2"]]
    
for tab_name in tab_order:
    pyautogui.hotkey('ctrl', 'alt', 'f5')
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'b')
    time.sleep(3)
    pyautogui.hotkey('ctrl', 'f4')
