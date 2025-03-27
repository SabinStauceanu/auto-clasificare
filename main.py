import time
import subprocess
import pyautogui
import xlwings as xw
from datetime import date
from datetime import datetime, timedelta

pyautogui.FAILSAFE = True

# Extragere date

wb = xw.Book("C:\\Users\\SAMI\\Desktop\\Document clasificare.xlsx").sheets['Sheet1']

lastCell = wb.range('F' + str(wb.cells.last_cell.row)).end('up').row

greutate = wb.range("F2:F" + str(lastCell)).value
calitate = wb.range("E2:E" + str(lastCell)).value
nrLot = wb.range("C2:C" + str(lastCell)).value
nrOrdine = wb.range("D2:D" + str(lastCell)).value
nrLuni = wb.range("B2:B" + str(lastCell)).value
sex_rasa = wb.range("G2:G" + str(lastCell)).value

# Introducere in P04

subprocess.run([r"\\tango1\\install\\vif\\viferptactile-prod.bat"])
time.sleep(10)
#pyautogui.moveTo(800, 350)
#pyautogui.click()
#pyautogui.moveTo(1141, 344)
#time.sleep(0.5)
#pyautogui.click()
pyautogui.moveTo(1253, 843)
time.sleep(0.5)
pyautogui.click()
pyautogui.moveTo(958, 649)
time.sleep(0.5)
pyautogui.click()
pyautogui.moveTo(935, 483)
time.sleep(0.5)
pyautogui.click()
pyautogui.moveTo(1247, 296)
time.sleep(0.5)
pyautogui.click()
pyautogui.moveTo(992, 603)
time.sleep(0.5)
pyautogui.click()

time.sleep(5)

for i in range(len(greutate)):

    pyautogui.moveTo(888, 374)
    pyautogui.click()
    pyautogui.moveTo(1400, 500)
    pyautogui.click()

    # Completare cantitate si numar carcase
    time.sleep(2)
    pyautogui.moveTo(1175, 620)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(950, 510)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(950, 750)
    time.sleep(1)
    pyautogui.click()
    pyautogui.moveTo(1180, 690)
    time.sleep(1)
    pyautogui.click()
    cantitate = list(str(int(greutate[i])))
    for j in range(len(cantitate)):
        if cantitate[j] == "1":
            pyautogui.moveTo(880, 500)
        elif cantitate[j] == "2":
            pyautogui.moveTo(950, 510)
        elif cantitate[j] == "3":
            pyautogui.moveTo(1020, 510)
        elif cantitate[j] == "4":
            pyautogui.moveTo(880, 450)
        elif cantitate[j] == "5":
            pyautogui.moveTo(950, 435)
        elif cantitate[j] == "6":
            pyautogui.moveTo(1020, 440)
        elif cantitate[j] == "7":
            pyautogui.moveTo(880, 373)
        elif cantitate[j] == "8":
            pyautogui.moveTo(950, 370)
        elif cantitate[j] == "9":
            pyautogui.moveTo(1020, 370)
        elif cantitate[j] == "0":
            pyautogui.moveTo(880, 580)
        pyautogui.click()
    time.sleep(1)
    pyautogui.moveTo(950, 750)
    pyautogui.click()
    pyautogui.moveTo(1241, 841)
    time.sleep(0.5)
    pyautogui.click()

    listCalitate = list(calitate[i])
    listNrLot = list(str(int(nrLot[i])))
    listNrOrdine = list(str(int(nrOrdine[i])))
    listNrLuni = list(str(int(nrLuni[i])))

    # Completare categorie carcasa
    pyautogui.moveTo(800, 315)
    time.sleep(3)
    pyautogui.click()
    time.sleep(1.5)
    for j in range(len(listCalitate)):
        if listCalitate[j] == "A":
            pyautogui.moveTo(700, 454)
        elif listCalitate[j] == "B":
            pyautogui.moveTo(765, 450)
        elif listCalitate[j] == "Z":
            pyautogui.moveTo(1110, 595)
        elif listCalitate[j] == "P":
            pyautogui.moveTo(1085, 520)
        elif listCalitate[j] == "V":
            pyautogui.moveTo(830, 595)
        elif listCalitate[j] == "O":
            pyautogui.moveTo(1015, 520)
        elif listCalitate[j] == "U":
            pyautogui.moveTo(765, 587)
        elif listCalitate[j] == "E":
            pyautogui.moveTo(945, 453)
        elif listCalitate[j] == "R":
            pyautogui.moveTo(1225, 520)
        elif listCalitate[j] == "D":
            pyautogui.moveTo(905, 451)
        elif listCalitate[j] == "+":
            pyautogui.moveTo(1285, 660)
            pyautogui.click()
            time.sleep(0.5)
            pyautogui.moveTo(835, 590)
            pyautogui.click()
            time.sleep(0.5)
            pyautogui.moveTo(1285, 660)
        elif listCalitate[j] == "-":
            pyautogui.moveTo(1188, 593)
        elif listCalitate[j] == "1":
            pyautogui.moveTo(667, 381)
        elif listCalitate[j] == "2":
            pyautogui.moveTo(735, 379)
        elif listCalitate[j] == "3":
            pyautogui.moveTo(807, 380)
        elif listCalitate[j] == "4":
            pyautogui.moveTo(877, 381)
        elif listCalitate[j] == "5":
            pyautogui.moveTo(944, 379)
        pyautogui.click()
    pyautogui.moveTo(1140, 770)
    time.sleep(0.5)
    pyautogui.click()

    # Completare numar lot
    pyautogui.moveTo(1009, 566)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(800, 315)
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(0.5)
    for j in range(len(listNrLot)):
        if listNrLot[j] == "1":
            pyautogui.moveTo(667, 381)
        elif listNrLot[j] == "2":
            pyautogui.moveTo(735, 379)
        elif listNrLot[j] == "3":
            pyautogui.moveTo(807, 380)
        elif listNrLot[j] == "4":
            pyautogui.moveTo(877, 381)
        elif listNrLot[j] == "5":
            pyautogui.moveTo(944, 379)
        elif listNrLot[j] == "6":
            pyautogui.moveTo(1014, 381)
        elif listNrLot[j] == "7":
            pyautogui.moveTo(1088, 381)
        elif listNrLot[j] == "8":
            pyautogui.moveTo(1154, 381)
        elif listNrLot[j] == "9":
            pyautogui.moveTo(1225, 374)
        elif listNrLot[j] == "0":
            pyautogui.moveTo(1295, 373)
        pyautogui.click()
    pyautogui.moveTo(1140, 770)
    time.sleep(0.5)
    pyautogui.click()

    # Completare numar de ordine
    pyautogui.moveTo(983, 470)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(800, 315)
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(0.5)
    for j in range(len(listNrOrdine)):
        if listNrOrdine[j] == "1":
            pyautogui.moveTo(667, 381)
        elif listNrOrdine[j] == "2":
            pyautogui.moveTo(735, 379)
        elif listNrOrdine[j] == "3":
            pyautogui.moveTo(807, 380)
        elif listNrOrdine[j] == "4":
            pyautogui.moveTo(877, 381)
        elif listNrOrdine[j] == "5":
            pyautogui.moveTo(944, 379)
        elif listNrOrdine[j] == "6":
            pyautogui.moveTo(1014, 381)
        elif listNrOrdine[j] == "7":
            pyautogui.moveTo(1088, 381)
        elif listNrOrdine[j] == "8":
            pyautogui.moveTo(1154, 381)
        elif listNrOrdine[j] == "9":
            pyautogui.moveTo(1225, 374)
        elif listNrOrdine[j] == "0":
            pyautogui.moveTo(1295, 373)
        pyautogui.click()
    pyautogui.moveTo(1140, 770)
    time.sleep(0.5)
    pyautogui.click()

    # completare varsta
    pyautogui.moveTo(1017, 623)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(800, 315)
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(0.5)
    for j in range(len(listNrLuni)):
        if listNrLuni[j] == "1":
            pyautogui.moveTo(667, 381)
        elif listNrLuni[j] == "2":
            pyautogui.moveTo(735, 379)
        elif listNrLuni[j] == "3":
            pyautogui.moveTo(807, 380)
        elif listNrLuni[j] == "4":
            pyautogui.moveTo(877, 381)
        elif listNrLuni[j] == "5":
            pyautogui.moveTo(944, 379)
        elif listNrLuni[j] == "6":
            pyautogui.moveTo(1014, 381)
        elif listNrLuni[j] == "7":
            pyautogui.moveTo(1088, 381)
        elif listNrLuni[j] == "8":
            pyautogui.moveTo(1154, 381)
        elif listNrLuni[j] == "9":
            pyautogui.moveTo(1225, 374)
        elif listNrLuni[j] == "0":
            pyautogui.moveTo(1295, 373)
        pyautogui.click()
    pyautogui.moveTo(1140, 770)
    time.sleep(0.5)
    pyautogui.click()

    # Selectare categorie pret
    pyautogui.moveTo(1025, 668)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(800, 315)
    time.sleep(1)
    pyautogui.click()
    if greautate[i] > 0 and greautate[i] <= 160 and nrLuni[i] >= 24:
        pyautogui.moveTo(600, 285)
    elif greautate[i] >= 161 and greautate[i] <= 200 and nrLuni[i] >= 24:
        pyautogui.moveTo(800, 285)
    elif greautate[i] >= 201 and greautate[i] <= 250 and nrLuni[i] >= 24:
        pyautogui.moveTo(1000, 285)
    elif greautate[i] >= 251 and greautate[i] <= 300 and nrLuni[i] >= 24:
        pyautogui.moveTo(1220, 285)
    elif greautate[i] >= 301 and greautate[i] <= 350 and nrLuni[i] >= 24:
        pyautogui.moveTo(600, 385)
    elif greautate[i] >= 351 and nrLuni[i] >= 24:
        pyautogui.moveTo(800, 385)
    elif sex_rasa[i] == "B":
        pyautogui.moveTo(1000, 385)
    elif nrLuni[i] < 24 and sex_rasa == "F":
        pyautogui.moveTo(600, 500)
    elif greautate[i] <= 200 and nrLuni[i] < 24 and sex_rasa == "M":
        pyautogui.moveTo(800, 500)
    elif greautate[i] >= 201 and nrLuni[i] < 24 and sex_rasa == "M":
        pyautogui.moveTo(1000, 500)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(1258, 840)
    time.sleep(0.5)
    pyautogui.click()

    # Iesire din fabricatie
    pyautogui.moveTo(1409, 591)
    time.sleep(1.2)
    pyautogui.click()
    pyautogui.moveTo(988, 616)
    time.sleep(1)
    pyautogui.click()
    pyautogui.moveTo(1404, 160)
    time.sleep(1)
    pyautogui.click()
    time.sleep(5)

pyautogui.hotkey("alt", "f4")
pyautogui.hotkey("alt", "f4")
