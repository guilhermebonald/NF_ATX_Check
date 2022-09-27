import pyautogui
from time import sleep
from PIL import Image
from pytesseract import pytesseract
from openpyxl import load_workbook
import pandas as pd
import win32print
import win32api

# Software para atualização do documento fiscal do programa ATX Frota.

# top and left: These parameters represent the top left coordinates i.e (x,y) = (left, top).
# bottom and right: These parameters represent the bottom right coordinates i.e. (x,y) = (right, bottom).

# use 'mouseinfo' to get coordinates x, y.

# set pytesseract path and instance it.

path_to_tesseract = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
pytesseract.tesseract_cmd = path_to_tesseract

# worksheet data set

wb = load_workbook(filename="data.xlsx")
sheet = wb.active

sheetNfListValue = []


# nf number list
nfList = []

# add sheet in list

nf = sheet["D"]
for i in nf:
    sheetNfListValue.append(i.value)


# set printer

impressoras = win32print.EnumPrinters(2)
brother = impressoras[6]
win32print.SetDefaultPrinter(brother[2])
arquivo = "result.xlsx"
caminho = r"C:\Users\fiscal-01.AGTURISMO\Resilio Sync\CODE\Bot\NFSCheck"

# check nf function


def check_nf(nf):
    sleep(2)
    # --- set nf number ---
    pyautogui.click(428, 233)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.press("backspace")
    pyautogui.write(nf)
    # --- filter nf and get screenshot image ---
    pyautogui.click(885, 201)
    pyautogui.screenshot("image.png", region=(395, 310, 50, 50))
    # --- image to text ---
    image = "image.png"
    text = pytesseract.image_to_string(Image.open(image))
    # --- validation ---
    if text == "":
        nf_number = 0
    else:
        nf_number = int(text)
    return nf_number


# loop list nf name


def generate_xlsx():
    # Append in List
    for g in sheetNfListValue:
        if check_nf(g) == 0:
            nfList.append(g)

    # Generate DataFrame
    d = {"NF NÃO LANÇADAS": nfList}
    df = pd.DataFrame(data=d)
    df.to_excel("result.xlsx")

    # imprimir
    win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)


generate_xlsx()
