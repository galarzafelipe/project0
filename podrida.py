#Podrida#

from openpyxl import Workbook

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
players = int(players)
playerList = []
for i in range(players):
    p = input("Please enter Player {}:".format(i + 1))
    playerList.append(p)
print(playerList)

###########OPCIONES###
####Que tipo de juego
##Reparte?

workbook = Workbook()
sheet = workbook.active

sheet.cell(row = 1, column = 1).value = "manos"
for x in range(1,maxcards + 1,1):
    sheet.cell(row = x + 1, column = 1).value = x
    sheet.cell(row = maxcards + players + x + 1, column = 1).value = abs(x - 11)
for y in range(1,players + 1,1):
    sheet.cell(row = 1, column = 2*y).value = playerList[y - 1]
    sheet.cell(row = y + maxcards + 1, column = 1).value = "{} ST".format(maxcards)

workbook.save(filename = "Trial1.xlsx")
print("Enjoy Playing!")
