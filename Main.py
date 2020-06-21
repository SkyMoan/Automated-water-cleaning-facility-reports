import xlsxwriter
import xlrd

loc = r"C:\Users\Erling\Desktop\service2020.xlsx"


def getSingleCell(row, col):
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    return(sheet.cell_value(row, col))



def getFullLine(row):
    x = 0
    for x in range(0,14):
        getSingleCell(row, x)


def getAllLines():
    for x in range(1,5):
        getFullLine(x)


def writeReport():


    workbook = xlsxwriter.Workbook('ServiceReport.xlsx')
    myworksheet = workbook.add_worksheet()

    #Formatting used

    header = workbook.add_format({'bold': True, 'font_size': '16'})
    bold = workbook.add_format({'bold': True})

    # Add default template
    myworksheet.write(5,1, "Serviceskjema Gråvannsrenseanlegg", header)
    myworksheet.write(8,1, "Kunde: ")
    myworksheet.write(9,1, "Anleggsadresse: ")
    myworksheet.write(12,1, "Anleggstype: ")
    myworksheet.write(13,1, "G. nr. og B. nr.: ")
    myworksheet.write(14,1, "Tømmestatus: ")
    myworksheet.write(16,1, "Slamavskiller", bold)
    myworksheet.write(18,1, "Størrelse")
    myworksheet.write(20,1, "Slamnivå i dag")
    myworksheet.write(22,1, "Slamnivå sist")
    myworksheet.write(24,1, "Pumpe 1", bold)
    myworksheet.write(26,1, "Rengjøring Vippefunksjoner")
    myworksheet.write(28,1, "Funksjon, lyd-/lydtest")
    myworksheet.write(30,1, "Signalgiver i kum")
    myworksheet.write(32,1, "Filterkum", bold)
    myworksheet.write(34,1, "Antall")
    myworksheet.write(35,1, "Størrelse ")
    myworksheet.write(36,1, "Dyser/Spredebilde")
    myworksheet.write(37,1, "Filterflate raking")
    myworksheet.write(40,1, "Andre merknader", bold)
    myworksheet.write(51,1, "Knut J Moen")
    myworksheet.write(52,1, "Signatur")

    myworksheet.write(16,4, "Merknad", bold)
    myworksheet.write(24,4, "Merknad", bold)
    myworksheet.write(16,9, "Merknad", bold)
    myworksheet.write(24,9, "Merknad", bold)
    myworksheet.write(32,4, "Merknad", bold)

    myworksheet.write(8,6, "Dato: ")
    myworksheet.write(9,6, "Forrige Service: ")
    myworksheet.write(12,6, "Antall PE:")
    myworksheet.write(14,6, "Lokksikring:")
    myworksheet.write(16,6, "Plassering kontrollboks/skap", bold)
    myworksheet.write(18,6, "Innendørs")
    myworksheet.write(20,6, "Ute på vegg")
    myworksheet.write(22,6, "Ute på stolpe")
    myworksheet.write(24,6, "Pumpe 2 til spredegrøft", bold)
    myworksheet.write(26,6, "Rengjøring, vippefunksjoner")
    myworksheet.write(28,6, "Funksjon, lyd-/lydtest")
    myworksheet.write(30,6, "Signalangiver i insp.kum/pumpekum")
    myworksheet.write(32,6, "Vannkvalitet", bold)
    myworksheet.write(34,6, "Lukt")
    myworksheet.write(36,6, "Farge")
    myworksheet.write(32,8, "Merknad", bold)

    x = 1

    #Adresse
    #myworksheet.write(#,# getSingleCell(x, 1)+ " "+str(getSingleCell(x,2)))

    #Slamnivå
    #myworksheet.write(#,#, getSingleCell(x, 4))

    #ServiceDato
    #myworksheet.write(#,#, getSingleCell(x,0))

    #Vannkvalitet
    #myworksheet.write(#,#, getSingleCell(x,8))
    #myworksheet.write(#,#, getSingleCell(x,12))

    #Lokksikring
    #myworksheet.write(#,#, getSingleCell(x,7))


    workbook.close()


writeReport()



