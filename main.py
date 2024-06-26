

import time
import subprocess
import pyautogui
import xlwings as xw
from datetime import date
from datetime import datetime, timedelta

pyautogui.FAILSAFE = True

#Extragere date

wb = xw.Book("C:\\Users\\CALITATE\\Desktop\\Document clasificare.xlsx").sheets['Sheet1']

lastCell = wb.range('G' + str(wb.cells.last_cell.row)).end('up').row

nrAbatorizare = wb.range("C2:C" + str(lastCell)).value
dataAbatorizare = wb.range("E2:E" + str(lastCell)).value
greutate = wb.range("G2:G" + str(lastCell)).value
calitate = wb.range("I2:I" + str(lastCell)).value
nrLot = wb.range("J2:J" + str(lastCell)).value
nrOrdine = wb.range("K2:K" + str(lastCell)).value
nrLuni = wb.range("L2:L" + str(lastCell)).value

today = date.today()
formatted_date = today.strftime('%d.%m.%Y')



#Introducere in P04

subprocess.run([r"\\tango1\\install\\vif\\viferptactile-prod.bat"])
time.sleep(5)
pyautogui.moveTo(800,350)
pyautogui.click()
pyautogui.moveTo(1141,344)
time.sleep(0.5)
pyautogui.click()
pyautogui.moveTo(1253,843)
time.sleep(0.5)
pyautogui.click()
pyautogui.moveTo(958,649)
time.sleep(0.5)
pyautogui.click()
pyautogui.moveTo(935,483)
time.sleep(0.5)
pyautogui.click()
pyautogui.moveTo(1247,296)
time.sleep(0.5)
pyautogui.click()
pyautogui.moveTo(992,603)
time.sleep(0.5)
pyautogui.click()

time.sleep(5)

#Selectare miscare finalizate
pyautogui.moveTo(1250, 300)
time.sleep(0.5)
pyautogui.click()
time.sleep(5)
pyautogui.moveTo(495, 288)
time.sleep(0.5)
pyautogui.click()
time.sleep(5)
pyautogui.moveTo(750, 289)
time.sleep(0.5)
pyautogui.click()
time.sleep(5)

data_curenta = datetime.strptime(formatted_date, "%d.%m.%Y")

nrReferintaClicks = 10

for i in range(len(greutate)):
    # Schimbare data
    data_tabel = datetime.strptime(dataAbatorizare[i].strftime('%d.%m.%Y'), "%d.%m.%Y")
    if (data_curenta - data_tabel).days > 0:
        pyautogui.moveTo(495, 288)
        time.sleep(0.5)
        pyautogui.click(clicks=(data_curenta - data_tabel).days, interval=5)
        data_curenta = data_tabel
        nrReferintaClicks = 10
    elif (data_curenta - data_tabel).days < 0:
        pyautogui.moveTo(750, 289)
        time.sleep(0.5)
        pyautogui.click(clicks=abs((data_curenta - data_tabel).days), interval=5)
        data_curenta = data_tabel
        nrReferintaClicks = 10


    #Selectare numar abatorizare
    if int(nrAbatorizare[i]) == 1:
        pyautogui.moveTo(985, 390)
    elif int(nrAbatorizare[i]) == 2:
        pyautogui.moveTo(890, 434)
    elif int(nrAbatorizare[i]) == 3:
        pyautogui.moveTo(900, 482)
    elif int(nrAbatorizare[i]) == 4:
        pyautogui.moveTo(911, 526)
    elif int(nrAbatorizare[i]) == 5:
        pyautogui.moveTo(889, 568)
    elif int(nrAbatorizare[i]) == 6:
        pyautogui.moveTo(884, 614)
    elif int(nrAbatorizare[i]) == 7:
        pyautogui.moveTo(900, 663)
    elif int(nrAbatorizare[i]) == 8:
        pyautogui.moveTo(900, 702)
    elif int(nrAbatorizare[i]) == 9:
        pyautogui.moveTo(900, 750)
    elif int(nrAbatorizare[i]) == 10:
        pyautogui.moveTo(900, 795)
    elif int(nrAbatorizare[i]) > 10:
        pyautogui.moveTo(1285, 780)
        pyautogui.click(clicks=int(nrAbatorizare[i])-int(nrReferintaClicks), interval=0.1)
        pyautogui.moveTo(900, 795)
        nrReferintaClicks = nrAbatorizare[i]
    pyautogui.click()
    pyautogui.moveTo(1400, 500)
    pyautogui.click()

    #Completare cantitate si numar carcase
    pyautogui.moveTo(1175, 620)
    time.sleep(0.5)
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
        time.sleep(0.5)
        pyautogui.click()
    pyautogui.moveTo(950, 750)
    pyautogui.click()
    pyautogui.moveTo(1180, 690)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(950, 510)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(950, 750)
    pyautogui.click()

    listCalitate = list(calitate[i])
    listNrLot = list(str(int(nrLot[i])))
    listNrOrdine = list(str(int(nrOrdine[i])))
    listNrLuni = list(str(int(nrLuni[i])))

    #Completare categorie carcasa
    pyautogui.moveTo(800, 315)
    time.sleep(0.5)
    pyautogui.click()
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
            pyautogui.moveTo(765, 975)
        elif listCalitate[j] == "E":
            pyautogui.moveTo(945, 453)
        elif listCalitate[j] == "R":
            pyautogui.moveTo(1225, 520)
        elif listCalitate[j] == "D":
            pyautogui.moveTo(905, 451)
        elif listCalitate[j] == "+":
            pyautogui.moveTo(1285, 660)
            time.sleep(0.5)
            pyautogui.click()
            pyautogui.moveTo(835, 590)
            time.sleep(0.5)
            pyautogui.click()
            pyautogui.moveTo(1285, 660)
            time.sleep(0.5)
            pyautogui.click()
            pyautogui.moveTo(1285, 660)
            time.sleep(0.5)
            pyautogui.click()
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
        time.sleep(0.5)
        pyautogui.click()
    pyautogui.moveTo(1140, 770)
    time.sleep(0.5)
    pyautogui.click()

    #Completare numar lot
    pyautogui.moveTo(1009,566)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(800, 315)
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(1)
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
            pyautogui.moveTo(1014,381)
        elif listNrLot[j] == "7":
            pyautogui.moveTo(1088,381)
        elif listNrLot[j] == "8":
            pyautogui.moveTo(1154,381)
        elif listNrLot[j] == "9":
            pyautogui.moveTo (1225,374)
        elif listNrLot[j] == "0":
            pyautogui.moveTo(1295,373)
        pyautogui.click()
    pyautogui.moveTo(1140, 770)
    time.sleep(0.5)
    pyautogui.click()

    #Completare numar de ordine
    pyautogui.moveTo(983,470)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(800, 315)
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(1)
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
            pyautogui.moveTo(1014,381)
        elif listNrOrdine[j] == "7":
            pyautogui.moveTo(1088,381)
        elif listNrOrdine[j] == "8":
            pyautogui.moveTo(1154,381)
        elif listNrOrdine[j] == "9":
            pyautogui.moveTo (1225,374)
        elif listNrOrdine[j] == "0":
            pyautogui.moveTo(1295,373)
        pyautogui.click()
    pyautogui.moveTo(1140, 770)
    time.sleep(0.5)
    pyautogui.click()

    #completare varsta
    pyautogui.moveTo(1017,623)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(800, 315)
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(1)
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
            pyautogui.moveTo(1014,381)
        elif listNrLuni[j] == "7":
            pyautogui.moveTo(1088,381)
        elif listNrLuni[j] == "8":
            pyautogui.moveTo(1154,381)
        elif listNrLuni[j] == "9":
            pyautogui.moveTo (1225,374)
        elif listNrLuni[j] == "0":
            pyautogui.moveTo(1295,373)
        pyautogui.click()
    pyautogui.moveTo(1140, 770)
    time.sleep(0.5)
    pyautogui.click()

    #Selectare categorie pret
    pyautogui.moveTo(1025,668)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(800, 315)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(1030,490)
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.moveTo(1258,840)
    time.sleep(0.5)
    pyautogui.click()

    #Iesire din fabricatie
    pyautogui.moveTo(1409,591)
    time.sleep(1)
    pyautogui.click()
    pyautogui.moveTo(988,616)
    time.sleep(1)
    pyautogui.click()
    pyautogui.moveTo(1404,160)
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(5)
