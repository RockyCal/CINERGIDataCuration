#checks for non-text symbols in resource titles, descriptions, and abstracts

import urllib.request
from openpyxl import load_workbook
from bs4 import BeautifulSoup

#excel file versions of inventories
senx = load_workbook("detailsSEN.xlsx")
sen = senx.active
senlinks = []

#cpx = load_workbook("detailsC4P.xlsx")
#cp = cpx.active
#cplinks = []

#hlx = load_workbook("detailsHL.xlsx")
#hl = hlx.active
#hllinks = []

#wsx = load_workbook("detailsWS.xlsx")
#ws = wsx.active
#wslinks = []

#inventories = [cp, hl, sen, ws]
#linklists = [cplinks, hllinks, senlinks, wslinks]

URL = "A"
start = 1

def list_of_links(workbook, llist):
    for row in workbook.range("%s%s:%s%s" % (URL, start, URL, workbook.get_highest_row())):
        for cell in row:
            llist.append(cell.value)

#returns list of resources with non-text symbols and the section that the symbol appears in
#if a location appears twice (e.g. <title>-abstract-abstract) it means that a non-text symbol
#appears twice in that location
def multiplefilechecker(llist):
    newlist = []
    for link in llist:
        soup = BeautifulSoup(urllib.request.urlopen(link))
        add = filechecker(soup)
        if len(add) > 1:
            newlist.append(add)
    return newlist

def filechecker(soup):
    fields = soup.find_all("td", limit = 4) #returns text of title, brief description, url, and abstract or purpose, respectively
    newstring = str(soup.find("td"))
    for item in fields:
        string = str(item)
        count = 0
        if fields.index(item) != 2:
            for char in string:
                if ord(char) > 126:
                    if fields.index(item) == 0:
                        newstring += "-title"
                    elif fields.index(item) == 1:
                        newstring += "-description"
                    elif fields.index(item) == 3:
                        newstring += "-abstract"
    if len(newstring) == len(str(soup.find("td"))):
        newstring = ""
    return newstring

list_of_links(sen, senlinks)
print(multiplefilechecker(senlinks))
