#Podrida#

from openpyxl import Workbook
from openpyxl.utils import get_column_letter, FORMULAE, column_index_from_string
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

#Cuantos jugadores?
#En que orden estan sentados?
#Sube de a 1 hasta el max y baja de a 1
#
#
##############
#players = input("How many players? ")
#players = int(players)
#zprint(players)

##############

players = input("How many players? ")
players = int(players)
maxcards = int(52 / players)
reductor = maxcards + 1
players = int(players)
playerList = []
evens = []
odd = []
numseparator = list(range(1, maxcards+1))
for num in numseparator:
    if num % 2 == 0:
        evens.append(num)
    else:
        odd.append(num)
print(evens)
print(odd)

for i in range(players):
    p = input("Please enter Player {}:".format(i + 1))
    playerList.append(p)
print(playerList)

##########OPCIONES###
####Que tipo de juego
##Reparte?
#Activate Workbook#
workbook = Workbook()
sheet = workbook.active
align = Alignment(horizontal = 'left', vertical = 'center')
pedidas = []
hechas = []
puntos = []
#Preparing Workbook#
for x in range(1,maxcards + 1,1):
    sheet.cell(row = x + 1, column = 1).value = x
    sheet.cell(row = x + 1, column = 1).alignment = align
    sheet.cell(row = maxcards + players + x + 1, column = 1).value = abs(x - reductor)
    sheet.cell(row = maxcards + players + x + 1, column = 1).alignment = align
for y in range(1,players + 1,1):
    sheet.cell(row = 1, column = 3*y-1).value = playerList[y - 1]
    sheet.cell(row = 1, column = 3*y-1).font = Font(bold = True)
    sheet.cell(row = 1, column = 3*y).value = "P?"
    cw = get_column_letter(3*y-1)
    cw1 = get_column_letter(3*y)
    sheet.column_dimensions[cw1].width = 4
    sheet.cell(row = 1, column = 3*y+1).value = "H?"
    cw2 = get_column_letter(3*y+1)
    sheet.column_dimensions[cw2].width = 4
    sheet.cell(row = y + maxcards + 1, column = 1).value = "{} ST".format(maxcards)
    sheet.cell(row = 2 * maxcards + players + 2, column = 1).value = "Puntos"
    pedidas.append(cw1)
    hechas.append(cw2)
    puntos.append(cw)


lastrow = 2 * maxcards + players + 2
lastcol = 3 * players + 2


if players == 5:
    for i in range(1, lastrow, 1):
        newi = str(i)
        n1 = pedidas[0] + newi
        n2 = pedidas[1] + newi
        n3 = pedidas[2] + newi
        n4 = pedidas[3] + newi
        n5 = pedidas[4] + newi
        m1 = hechas[0] + newi
        m2 = hechas[1] + newi
        m3 = hechas[2] + newi
        m4 = hechas[3] + newi
        m5 = hechas[4] + newi

        sheet.cell(row = i, column = lastcol).value = "=SUM({},{},{},{},{})".format(n1,n2,n3,n4,n5)
        sheet.cell(row = i, column = lastcol+1).value = "=SUM({},{},{},{},{})".format(m1,m2,m3,m4,m5)

#cip = []
#cih = []
#cipu = []
#for i in range(len(pedidas)):
#    x = pedidas[i]
#    y = column_index_from_string(x)
#    cip.append(y)
#for i in range(len(hechas)):
#    x = hechas[i]
#    y = column_index_from_string(x)
#    cih.append(y)
#for i in range(len(puntos)):
#    x = puntos[i]
#    y = column_index_from_string(x)
#    cipu.append(y)
#print(cip)
#print(cih)
#print(cipu)


sheet.cell(row = 1, column = lastcol).value = "Basas Pedidas"
sheet.cell(row = 1, column = lastcol +1).value = "Basas Completas"


#BORDERS#

thin_border = Border(left=Side(style='none'),
                     right=Side(style='thin'),
                     top=Side(style='none'),
                     bottom=Side(style='none'))

thick_border = Border(left=Side(style='none'),
                     right=Side(style='thick'),
                     top=Side(style='none'),
                     bottom=Side(style='none'))

thick_border1 = Border(left=Side(style='none'),
                     right=Side(style='thick'),
                     top=Side(style='thick'),
                     bottom=Side(style='thick'))

thin_border2 = Border(left=Side(style='none'),
                     right=Side(style='thin'),
                     top=Side(style='thick'),
                     bottom=Side(style='thick'))


for z in range(1,lastrow+1,1):
    for j in range(1, players + 1, 1):
            sheet.cell(row = z, column = 3 * j - 1).border = thin_border
            sheet.cell(row = z, column = 3 * j).border = thin_border
            sheet.cell(row = z, column = 3 * j + 1).border = thick_border
            sheet.cell(row = z, column = lastcol).border = thick_border1
            sheet.cell(row = z, column = lastcol + 1).border = thick_border1
    sheet.cell(row = z, column = 1).border = thick_border

for j in range(1, players + 1, 1):
    sheet.cell(row = lastrow, column = 3 * j - 1).border = thin_border2
    sheet.cell(row = lastrow, column = 3 * j).border = thin_border2
    sheet.cell(row = lastrow, column = 3 * j + 1).border = thick_border1
sheet.cell(row = lastrow, column = 1).border = thick_border1

#    sheet.cell(row = lastrow - 1, column = i).border = thick_border1
#    sheet.cell(row = lastrow, column = i).border = thick_border1
#BORDERS#


##FUNCTIONALITY###
























#Preparing Workbook#

workbook.save(filename = "Trial1.xlsx")
print("Enjoy Playing!")
