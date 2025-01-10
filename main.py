

import time
import subprocess
import pyautogui
import xlwings as xw
from datetime import date
from datetime import datetime, timedelta

pyautogui.FAILSAFE = True

#Extragere date

wb = xw.Book("C:\\Users\\SAMI\\Desktop\\Document clasificare.xlsx").sheets['Sheet1']

lastCell = wb.range('H' + str(wb.cells.last_cell.row)).end('up').row

nrAbatorizare = wb.range("H2:H" + str(lastCell)).value
dataAbatorizare = wb.range("D2:D" + str(lastCell)).value
greutate = wb.range("G2:G" + str(lastCell)).value
calitate = wb.range("F2:F" + str(lastCell)).value
nrLot = wb.range("C2:C" + str(lastCell)).value
nrOrdine = wb.range("E2:E" + str(lastCell)).value
nrLuni = wb.range("B2:B" + str(lastCell)).value
#cantitate = wb.range("C2:C" + str(lastCell)).value

today = date.today()
formatted_date = today.strftime('%d.%m.%Y')

#Introducere in P04

subprocess.run([r"\\tango1\\install\\vif\\viferptactile-prod.bat"])
time.sleep(6)
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
time.sleep(1)
pyautogui.click()
time.sleep(5)

data_curenta = datetime.strptime(formatted_date, "%d.%m.%Y")

nrReferintaMin = 1
nrReferinta1 = 1
nrReferinta2 = 2
nrReferinta3 = 3
nrReferinta4 = 4
nrReferinta5 = 5
nrReferinta6 = 6
nrReferinta7 = 7
nrReferinta8 = 8
nrReferinta9 = 9
nrReferinta10 = 10
nrReferintaMax = 10
dejaIntroduse = [""]


for i in range(len(greutate)):
    # Se selecteaza cortalul daca numarul la categorie cantitate este 4

    # Schimbare data
    data_tabel = datetime.strptime(dataAbatorizare[i], "%d.%m.%Y")
    if (data_curenta - data_tabel).days > 0:
        pyautogui.moveTo(495, 288)
        time.sleep(1)
        pyautogui.click(clicks=(data_curenta - data_tabel).days, interval=5)
        data_curenta = data_tabel
        nrReferintaMin = 1
        nrReferinta1 = 1
        nrReferinta2 = 2
        nrReferinta3 = 3
        nrReferinta4 = 4
        nrReferinta5 = 5
        nrReferinta6 = 6
        nrReferinta7 = 7
        nrReferinta8 = 8
        nrReferinta9 = 9
        nrReferinta10 = 10
        nrReferintaMax = 10
    elif (data_curenta - data_tabel).days < 0:
        pyautogui.moveTo(750, 289)
        time.sleep(0.5)
        pyautogui.click(clicks=abs((data_curenta - data_tabel).days), interval=5)
        data_curenta = data_tabel
        nrReferintaMin = 1
        nrReferinta1 = 1
        nrReferinta2 = 2
        nrReferinta3 = 3
        nrReferinta4 = 4
        nrReferinta5 = 5
        nrReferinta6 = 6
        nrReferinta7 = 7
        nrReferinta8 = 8
        nrReferinta9 = 9
        nrReferinta10 = 10
        nrReferintaMax = 10

    # CLICK NECESAR IN CAZUL IN CARE SE SELECTEAZA A 10-A POZITIE SAU MAI JOS
    necesaryClick = 0

    #Selectare numar abatorizare
    time.sleep(1)
    if int(nrAbatorizare[i]) == nrReferinta1:
        pyautogui.moveTo(888,374)
    elif int(nrAbatorizare[i]) == nrReferinta2:
        pyautogui.moveTo(890, 434)
    elif int(nrAbatorizare[i]) == nrReferinta3:
        pyautogui.moveTo(900, 482)
    elif int(nrAbatorizare[i]) == nrReferinta4:
        pyautogui.moveTo(911, 526)
    elif int(nrAbatorizare[i]) == nrReferinta5:
        pyautogui.moveTo(889, 568)
    elif int(nrAbatorizare[i]) == nrReferinta6:
        pyautogui.moveTo(884, 614)
    elif int(nrAbatorizare[i]) == nrReferinta7:
        pyautogui.moveTo(900, 663)
    elif int(nrAbatorizare[i]) == nrReferinta8:
        pyautogui.moveTo(900, 702)
    elif int(nrAbatorizare[i]) == nrReferinta9:
        pyautogui.moveTo(900, 750)
    elif int(nrAbatorizare[i]) == nrReferinta10:
        pyautogui.moveTo(879, 797)
        necesaryClick = 1
    elif int(nrAbatorizare[i]) > nrReferintaMax:
        pyautogui.moveTo(1285, 780)
        pyautogui.click(clicks=int(nrAbatorizare[i])-int(nrReferintaMax), interval=0.1)
        pyautogui.moveTo(879, 797)
        nrReferintaMin = nrReferintaMin + (int(nrAbatorizare[i])-int(nrReferintaMax))
        nrReferinta1 = nrReferinta1 + (int(nrAbatorizare[i]) - int(nrReferintaMax))
        nrReferinta2 = nrReferinta2 + (int(nrAbatorizare[i]) - int(nrReferintaMax))
        nrReferinta3 = nrReferinta3 + (int(nrAbatorizare[i]) - int(nrReferintaMax))
        nrReferinta4 = nrReferinta4 + (int(nrAbatorizare[i]) - int(nrReferintaMax))
        nrReferinta5 = nrReferinta5 + (int(nrAbatorizare[i]) - int(nrReferintaMax))
        nrReferinta6 = nrReferinta6 + (int(nrAbatorizare[i]) - int(nrReferintaMax))
        nrReferinta7 = nrReferinta7 + (int(nrAbatorizare[i]) - int(nrReferintaMax))
        nrReferinta8 = nrReferinta8 + (int(nrAbatorizare[i]) - int(nrReferintaMax))
        nrReferinta9 = nrReferinta9 + (int(nrAbatorizare[i]) - int(nrReferintaMax))
        nrReferinta10 = nrReferinta10 + (int(nrAbatorizare[i]) - int(nrReferintaMax))
        nrReferintaMax = nrAbatorizare[i]
        necesaryClick = 1
    elif int(nrAbatorizare[i]) < nrReferintaMin:
        pyautogui.moveTo(1287, 396)
        pyautogui.click(clicks= int(nrReferintaMin) - int(nrAbatorizare[i]), interval=0.1)
        pyautogui.moveTo(888,374)
        nrReferintaMax = nrReferintaMax - (int(nrReferintaMin) - int(nrAbatorizare[i]))
        nrReferinta1 = nrReferinta1 - (int(nrReferintaMin) - int(nrAbatorizare[i]))
        nrReferinta2 = nrReferinta2 - (int(nrReferintaMin) - int(nrAbatorizare[i]))
        nrReferinta3 = nrReferinta3 - (int(nrReferintaMin) - int(nrAbatorizare[i]))
        nrReferinta4 = nrReferinta4 - (int(nrReferintaMin) - int(nrAbatorizare[i]))
        nrReferinta5 = nrReferinta5 - (int(nrReferintaMin) - int(nrAbatorizare[i]))
        nrReferinta6 = nrReferinta6 - (int(nrReferintaMin) - int(nrAbatorizare[i]))
        nrReferinta7 = nrReferinta7 - (int(nrReferintaMin) - int(nrAbatorizare[i]))
        nrReferinta8 = nrReferinta8 - (int(nrReferintaMin) - int(nrAbatorizare[i]))
        nrReferinta9 = nrReferinta9 - (int(nrReferintaMin) - int(nrAbatorizare[i]))
        nrReferinta10 = nrReferinta10 - (int(nrReferintaMin) - int(nrAbatorizare[i]))
        nrReferintaMin = nrAbatorizare[i]

    # Scoatere din fabricatie daca a fost introduse

    #if len(dejaIntroduse) > 0:
    #    for j in range(len(dejaIntroduse)):
    #        if nrAbatorizare[i] == dejaIntroduse[j]:
    #            pyautogui.moveTo(1397, 596)
    #            time.sleep(0.5)
    #            pyautogui.click()
    #            pyautogui.moveTo(989, 621)
    #            time.sleep(0.5)
    #            pyautogui.click()
    #            time.sleep(5)



    pyautogui.click()
    pyautogui.moveTo(1400, 500)
    pyautogui.click()

    #Completare cantitate si numar carcase
    time.sleep(1)
    pyautogui.moveTo(1175, 620)
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(1)
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
    time.sleep(1.5)
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

    #Completare numar lot
    pyautogui.moveTo(1009,566)
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
    time.sleep(1.2)
    pyautogui.click()
    pyautogui.moveTo(988,616)
    time.sleep(1)
    pyautogui.click()
    pyautogui.moveTo(1404,160)
    time.sleep(1)
    pyautogui.click()
    time.sleep(5)

    #Adaugare in lista de introduse
    dejaIntroduse.append(nrAbatorizare[i])

    if necesaryClick == 1:
        pyautogui.moveTo(1287, 396)
        time.sleep(0.5)
        pyautogui.click()

pyautogui.hotkey("alt","f4")
pyautogui.hotkey("alt","f4")
