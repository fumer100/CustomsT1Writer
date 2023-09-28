import datetime
from collections import Counter

import pandas as pd


def createT1(file):
    # read xls ,save specific columns.
    with pd.ExcelFile(file) as xls:
        packList = pd.read_excel(xls, 'Packing List')
        civ = pd.read_excel(xls, 'Commercial Invoice')

    startIndexPacklist = 7
    startIndexCiv = 8

    civLen, packLen = getLengthOfColumns(packList, civ)
    Itemlist = packList[startIndexPacklist:packLen]['Unnamed: 2'].reset_index(drop=True)  # ready
    CTN = packList[startIndexPacklist:packLen]['Unnamed: 6'].reset_index(drop=True)  # ready
    GrossWeight = packList[startIndexPacklist:packLen]['Unnamed: 9'].reset_index(drop=True)  # ready
    NetWeight = packList[startIndexPacklist:packLen]['Unnamed: 11'].reset_index(drop=True)  # ready
    CustomsValue3 = civ[startIndexCiv:7 + civLen]['Unnamed: 6'].reset_index(drop=True)  # ready
    hsCode = civ[startIndexCiv:7 + civLen]['Unnamed: 9'].reset_index(drop=True)  # ready
    rechnungNr = civ.iloc[4]['Unnamed: 4']
    rechnungsDatum=civ.iloc[5]['Unnamed: 4']
    containerNr=civ.iloc[5]['Unnamed: 1']

    # Erstellung der Tabelle

    dict = {'AT/B Position': "", 'H.S code': hsCode, "Ursprung": "CN", 'Warenbeschreibung': Itemlist,
            'Packstückanzahl': CTN, 'Gewicht': GrossWeight,
            'Netto': NetWeight, 'Warenwert': CustomsValue3, "Besonderheit": "ydocs"}  # y-docs fehlen mit macro
    df = pd.DataFrame(dict)

    # Erstellung der T1
    T1 = df.reset_index(drop=True)
    print(T1)

    # Erstellung der Verzollung
    verzollung = createVerzollung(T1)
    print(type(verzollung))
    T1["Warenwert"] = [round(a, 2) for a in T1["Warenwert"]]

    T1 = addYcodes(T1)
    T1 = addSums(T1)
    return T1, verzollung, rechnungNr, rechnungsDatum, containerNr


def remove_duplicates(input):
    # split input string separated by space
    input = input.split(",\t")

    # now create dictionary using counter method
    # which will have strings as key and their
    # frequencies as value
    UniqW = Counter(input)
    # joins two adjacent elements in iterable way
    s = " ".join(UniqW.keys())
    return s

def writeToExcel(writer, verzollung, T1, Codes):  # DONT TOUCH
    verzollung.to_excel(writer, sheet_name='verzollung', startrow=9, startcol=0, index=False)
    T1.to_excel(writer, sheet_name='T1', startrow=9, startcol=0, index=False)
    Codes.to_excel(writer, sheet_name='Codes', startrow=0, startcol=0, index=False)

    return

def createVerzollung(T1):  # DONT TOUCH #FINAL
    join_unique = lambda x: ','.join(set(x))
    leadingSpaces = lambda x: x.lstrip()
    unique = lambda x: remove_duplicates(x)

    warenbesch1 = T1.groupby('H.S code').agg({'Warenbeschreibung': join_unique})  # erzeugt Warenbeschreibung für jede HS
    warenbesch1 = warenbesch1.agg({'Warenbeschreibung': leadingSpaces})
    warenbesch1 = warenbesch1.agg({'Warenbeschreibung': unique})

    verzollung = T1.groupby(['H.S code']).sum()  # Summiert alles für jede HS
    verzollung["Warenwert"] = [round(a, 2) for a in verzollung["Warenwert"]]

    verzollung['Warenbeschreibung'] = warenbesch1['Warenbeschreibung']  # ersetzt Warenbeschreibung mit richtiger.
    verzollung['Ursprung'] = ["CN" for i in verzollung['Ursprung']]
    verzollung['Besonderheit'] = ["y-docs" for i in verzollung['Besonderheit']]  # Change to y/code Macro later
    verzollung = verzollung.reset_index()
    verzollung = verzollung.set_index('AT/B Position')
    # print("Atb als index",verzollung)
    verzollung['index_column'] = verzollung.index
    # print("Indexspalte geaddet ",verzollung)
    verzollung = verzollung.reset_index()
    verzollung = verzollung.set_index('index_column')
    # print("indexspalte als neuer index",verzollung)
    verzollung = addYcodes(verzollung)
    verzollung = addSums(verzollung)

    print(verzollung)
    return verzollung


def addSums(table):  # dont Touch
    summs = ["", "", "", "Summe", sum(table['Packstückanzahl']), sum(table['Gewicht']), sum(table['Netto']),
             sum(table['Warenwert']), ""]
    table.loc[len(table)] = ["", "", "", '', '', '', '', "", ""]
    table.loc[len(table) + 1] = ["", "", "", '', '', '', '', "", ""]
    table.loc[len(table) + 2] = summs

    return table


def addYcodes(table):
    table['Besonderheit'] = ["=SVERWEIS(B" + str(i) + "&C" + str(
        i) + ";'https://clasquinsa.sharepoint.com/sites/CLQ-DUS/Shared Documents/General/DUS/Seefracht/Import Kunden/Yahee/Austarifierung/SDUS012030.xlsx'#$Codes.A$1:C$1048576;3;0)"
                             for i in range(11, 11 + len(table))]

    return table


def getLengthOfColumns(packList, civ):  # DONT TOUCH #FINAL
    lengthCiv = len(civ['Unnamed: 9'].dropna())
    lengthPackList = len(packList['Unnamed: 6']) - 1
    return lengthCiv, lengthPackList


def createWorkbook(T1, verzollung, Sendungsnr, Schiff, BL, BLDatum, Rechnungsnr, RechnungsDatum, Containernr, Incoterm,
                   Transportpreis, Inlandpreis, writer):
    with pd.ExcelFile("Y-Docs.xls") as xls:
        codes = pd.read_excel(xls, 'Codes')
    writeToExcel(writer, verzollung, T1, codes)
    print('verzollung', verzollung)

    workbook =  writer.book
    date_format = workbook.add_format({'num_format': 'dd.mm.yy'})

    T1Sheet = writer.sheets['T1']
    VerzollungsSheet = writer.sheets['verzollung']
    T1Sheet.set_column(0, 10, 25)
    VerzollungsSheet.set_column(0, 10, 25)

    T1Sheet.write('A1', 'Sendungsnummer:')
    T1Sheet.write('B1', Sendungsnr)

    T1Sheet.write('A2', 'Schiff:')
    T1Sheet.write('B2', Schiff)

    T1Sheet.write('A3', 'B/L Nummer:')
    T1Sheet.write('B3', BL)
    T1Sheet.write('C3', BLDatum, date_format)

    T1Sheet.write('A4', 'Rechnung:')
    T1Sheet.write('B4', Rechnungsnr)
    print(type(RechnungsDatum))
    T1Sheet.write_datetime('C4', RechnungsDatum, date_format)

    T1Sheet.write('A5', 'Container:')
    T1Sheet.write('B5', Containernr)

    T1Sheet.write('A6', 'Incoterm:')
    T1Sheet.write('B6', Incoterm)

    T1Sheet.write('A7', 'Transportpreis:')
    T1Sheet.write('B7', Transportpreis)

    T1Sheet.write('A8', 'Inland:')
    T1Sheet.write('B8', Inlandpreis)

    VerzollungsSheet.write('A1', 'Sendungsnummer:')
    VerzollungsSheet.write('B1', Sendungsnr)

    VerzollungsSheet.write('A2', 'Schiff:')
    VerzollungsSheet.write('B2', Schiff)

    VerzollungsSheet.write('A3', 'B/L Nummer:')
    VerzollungsSheet.write('B3', BL)
    VerzollungsSheet.write('C3', BLDatum)

    VerzollungsSheet.write('A4', 'Rechnung:')
    VerzollungsSheet.write('B4', Rechnungsnr)
    VerzollungsSheet.write_datetime('C4', RechnungsDatum, date_format)

    VerzollungsSheet.write('A5', 'Container:')
    VerzollungsSheet.write('B5', Containernr)
    VerzollungsSheet.write('A6', 'Incoterm:')
    VerzollungsSheet.write('B6', Incoterm)

    VerzollungsSheet.write('A7', 'Transportpreis:')
    VerzollungsSheet.write('B7', Transportpreis)

    VerzollungsSheet.write('A8', 'Inland:')
    VerzollungsSheet.write('B8', Inlandpreis)
    writer.close()
    return
