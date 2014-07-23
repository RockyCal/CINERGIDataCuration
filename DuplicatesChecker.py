#attempt at writing a program to check for duplicates among the EC inventories
from openpyxl import load_workbook, cell

wsx = load_workbook("ECWSRI.xlsx")
ws = wsx.active

hlx = load_workbook("ECHLRI.xlsx")
hl = hlx.active

senx = load_workbook("ECSENRI.xlsx")
sen = senx.active

cpx = load_workbook("ECC4PRI.xlsx")
cp = cpx.active

TITLE = "A"
start = 2

def list_of_titles(workbook, tlist):
    for row in workbook.range("%s%s:%s%s"%(TITLE, start, TITLE, workbook.get_highest_row())):
        for cell in row:
            tlist.append(cell.value)

wst = []
hlt = []
sent = []
cpt = []            
            
list_of_titles(ws, wst)
list_of_titles(hl, hlt)
list_of_titles(sen, sent)
list_of_titles(cp, cpt)

#to check to see if list_of_titles() works
print(len(wst))
print(len(hlt))
print(len(sent))
print(len(cpt))

def dupcheck(list1, list2, list3):
    for value in list1:
        if value in list2 and value not in list3:
            list3.append(value)

duplicates = []

#need to find a more efficient way to go through lists, or make a more efficient method to do this
#also need to figure out how to indicate which lists the duplicates show up in
dupcheck(sent, wst, duplicates)
dupcheck(sent, hlt, duplicates)
dupcheck(sent, cpt, duplicates)
dupcheck(wst, hlt, duplicates)
dupcheck(wst, cpt, duplicates)
dupcheck(hlt, cpt, duplicates)

print(duplicates)
print("done")

