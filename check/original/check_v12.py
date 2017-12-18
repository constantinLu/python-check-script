# -*- coding: latin-1' -*-

import os, arcpy
from win32com.client import Dispatch
from collections import OrderedDict
import xlsxwriter
import sys

#default encoding for xlsxWriter
reload(sys)
sys.setdefaultencoding('latin-1')  #utf-8

arcpy.env.overwriteOutput = True
# workspace=arcpy.GetParameterAsText(0)

#method for rgb color - win32.client
def rgb_to_hex(rgb):
    strValue = '%02x%02x%02x' % rgb
    iValue = int(strValue, 16)
    return iValue

#method for checking the and adding all the errors to one single cell
def checkError(dictEr, nameField, typeEr, valEr):
    if typeEr not in dictEr[nameField]:
        dictEr[nameField][typeEr]=[]
    if valEr not in dictEr[nameField][typeEr]:
        dictEr[nameField][typeEr].append(valEr)




# #path of the input
# inputPath=r"C:\Users\LunguC\Desktop\check\20171212_QA-QC"
# tempxlsPath = r"C:\Users\LunguC\Desktop\check\TemplateReport.xlsx"
# outPath= r"C:\Users\LunguC\Desktop\check\20171212_REPORT\TEST"


# testing
inputPath = r"C:\Users\LunguC\Desktop\test\TESTOUT"
tempxlsPath = r"C:\Users\LunguC\Desktop\test\TemplateReport.xlsx"
outPath = r"C:\Users\LunguC\Desktop\test\Output Excel"


#env
arcpy.env.workspace = inputPath

#creating a fileGeodatabase
arcpy.CreateFileGDB_management(outPath, 'checkGeometry.gdb')
gdbPath=os.path.join(outPath,'checkGeometry.gdb')

# dict for checkProblems table
d_problems = {"self intersections": "Selbstüberschneidung",
              "short segments": "Kurze segmente",
              "null geometry": "Objekte mit leerer Geometrie",
              "incorrect ring ordering": "Falsche Ausrichtung von Polygonen",
              "unclosed rings": "Offene Polygone / Polylinien",
              "empty parts": "Objekte ohne Geometrie",
              "duplicate vertex": "Doppelte Stützpunkte",
              "mismatched attributes": "Fehlerhafte Z- oder M-Koordinaten",
              "discontinuous parts": "Unregelmäßige Abschnitte",
              "empty Z values": "Leere Z-Werte",
              "bad envelope": "Falsche räumliche Ausdehnung von Objekten",
              "bad dataset extent": "Falsche räumliche Ausdehnung des Datensatzes",
              "Field not existent" : "Feld existiert nicht"
              }

# dict for adding data to column 3
d_best = {
    "O_SHAPE": "Originalname des Basisdaten-Shapes",
    "O_STAND": "Datenstand",
    "O_DATUM": "Datum zum Datenstand",
    "O_HERKUNFT": "Firmenname des Bearbeiters",
    "U_SHAPE": "Name des Shapes der Basisbearbeitungsstufe",
    "U_BEARB": "Firmenname des Bearbeiters",
    "BEMERKUN": "zusätzliche Informationen",
    "BL": "2stelliges Bundeslandkürzel gemäß ISO 3166-2, maßgeblich ist die herausgebende Stelle, bei deutschlandweiten Daten DE ",
    "GEBIET": "Gebietsabgrenzung, entweder des RBZ, Landkreises, Planungsgemeinschaft, o.ä.",
    "KATEGORIE": "Vorranggebiet, Schutzgebiet, Datenart o.ä.",
    "PROJEKT": "Kurzbezeichnung des gesamten Projektes",
    "ABSCHNITT": "Abschnitt, falls sinnvoll",
    "NAME": "Name oder sinnvoller Identifikator des Objektes (Schutzgebietes o.ä.)",
    "ZONE": "NUR BEI SCHUTZGEBIETEN! Zone zum Schutzgebiet (Ziffern verwenden, betrifft nur WSG, BSR)",
    "EU_NR": "NUR BEI SCHUTZGEBIETEN! EU-Meldnummer für Natura 2000 (mit 'DE' und ohne Leerzeichen)"
}

#dict template to compare the data
# d_reqLower = OrderedDict()
# d_reqLower["O_Shape"] = ["String", 100]
# d_reqLower["O_Stand"] = ["String", 100]
# d_reqLower["O_Datum"] = ["Date"]
# d_reqLower["O_Herkunft"] = ["String", 100]
# d_reqLower["U_Shape"] = ["String", 100]
# d_reqLower["U_Bearb"] = ["String", 50, "TNL", "IBUe", "Giftge", "Arcadis"]
# d_reqLower["Bemerkun"] = ["String", 250]
# d_reqLower["BL"] = ["String", 5, "BY", "ST", "SN", "TH", "DE"]
# d_reqLower["Gebiet"] = ["String", 100]
# d_reqLower["Kategorie"] = ["String", 100]
# d_reqLower["Projekt"] =["String", 20, "SuedOstLink"]
# d_reqLower["Abschnitt"] =  ["String", 10]
# d_reqLower["Name"] = ["String", 250]
# d_reqLower["Zone"] =  ["String", 100]
# d_reqLower["EU_Nr"] =  ["String", 20]

d_req = OrderedDict()
d_req["O_SHAPE"] = ["String", 100]
d_req["O_STAND"] = ["String", 100]
d_req["O_DATUM"] = ["Date"]
d_req["O_HERKUNFT"] = ["String", 100]
d_req["U_SHAPE"] = ["String", 100]
d_req["U_BEARB"] = ["String", 50, "TNL", "IBUe", "Giftge", "Arcadis"]
d_req["BEMERKUN"] = ["String", 250]
d_req["BL"] = ["String", 5, "BY", "ST", "SN", "TH", "DE"]
d_req["GEBIET"] = ["String", 100]
d_req["KATEGORIE"] = ["String", 100]
d_req["PROJEKT"] =["String", 20, "SuedOstLink"]
d_req["ABSCHNITT"] =  ["String", 10]
d_req["NAME"] = ["String", 250]
d_req["ZONE"] =  ["String", 100]
d_req["EU_NR"] =  ["String", 20]




#da.Walk - looking for FeatureClasses
for dirpath, dirnames, filenames in arcpy.da.Walk(inputPath, "FeatureClass"):
    for filename in filenames:

        print "-------------------------------STARTING -------------------------:", filename
        filename_path = os.path.join(dirpath, filename)  # set the path of the input shapefile/featureclass
        # print filename_path
        fieldList_brut = [i for i in arcpy.ListFields(filename_path)]


        #extracting the factoryCode for each shapFile
        code = arcpy.Describe(filename_path).spatialReference.factoryCode


        '''########################################
        EXCEL PART: Formating the cells (XLSXWRITER)
        ############################################'''

        #creating filenames for xls file
        xlName = os.path.splitext(filename)[0] + ".xlsx"  # the name of the output excel
        xlsPath = os.path.join(outPath, xlName)  # the path of the output excel

        #xlsxWriter (creating the file and the workSheets
        wb = xlsxwriter.Workbook(xlsPath.encode('latin-1'))

        #checkWorkbook
        ws = wb.add_worksheet(u'Prüfergebnis Pflichtattribute')
        ws.set_landscape() #LANDSCAP
        ws.set_paper(9) #A4
        ws.set_footer(filename) #footer
        ws.autofilter('A1:G1')

        #geometryWorkbook
        wsG = wb.add_worksheet(u'Prüfergebnis Geometrie')
        wsG.set_landscape() #LANDSCAP
        wsG.set_paper(9) #A4
        wsG.set_footer(filename) #footer
        wsG.autofilter('A1:D1')

        # defining parameters for each colon
        nameCell = 0
        typeCell = 1
        beschreibungCell = 2
        possibleCell = 3
        fieldCell = 4
        attCell = 5
        remarksCell = 6

        # set the dimmensions of the colons and setting some default colors
        ws.set_column(nameCell, nameCell, 10)
        ws.set_column(typeCell, typeCell, 13)
        ws.set_column(beschreibungCell, beschreibungCell, 40)
        ws.set_column(possibleCell, possibleCell, 20)
        ws.set_column(fieldCell, fieldCell, 16)
        ws.set_column(attCell, attCell, 100)
        ws.set_column(remarksCell, remarksCell, 40)


        #adding custom colors
        #adding the wraping to the larger columns (attributCells and possibleCells)
        style = wb.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 10, 'font_color': 'black', 'align' : 'left'})
        style.set_text_wrap()

        generalStyle = wb.add_format({'font_name': 'Arial', 'font_size': 10, 'align': 'vcenter'})
        generalStyle.set_text_wrap()

        #adding formatting styles for required cells
        #green = easyxf('pattern: pattern solid,fore_colour green_custom;' 'font: color green;' 'align: vert center')
        green = wb.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 10, 'font_color': 'green', 'bg_color': '#C6EFCD', 'align': 'vcenter'})
        green.set_text_wrap()

        grey = wb.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10, 'font_color': 'white', 'bg_color': '#A5A5A5'})
        grey.set_text_wrap()

        pink = wb.add_format({'bold': False, 'font_name': 'Arial', 'font_size': 10, 'font_color': '#ae5464', 'bg_color': '#FFC6CE', 'align': 'vcenter'})
        pink.set_text_wrap()


        '''----------ATTRIBUTE SHEET------'''
        #adding headers
        ws.write(0, nameCell, "Attribut", grey)
        ws.write(0, typeCell, "Parameter", grey)
        ws.write(0, beschreibungCell, "Beschreibung", grey)
        ws.write(0, possibleCell, "Mögliche Inhalte", grey)
        ws.write(0, fieldCell, 'Prüfergebnis', grey)
        ws.write(0, attCell, "Tatsächliche Inhalte", grey)
        ws.write(0, remarksCell, "Bemerkung", grey)

        '''----------GEOMETRY SHEET------'''
        #adding headers for second sheet
        wsG.write(0, 0, "Laufende Nummer", grey)
        wsG.write(0, 1, "Objektklasse", grey)
        wsG.write(0, 2, "Objekt-ID", grey)
        wsG.write(0, 3, "Problem", grey)

        #set dim for columns of -----GEOMETRY SHEET-----
        wsG.set_column(0, 0, 12)
        wsG.set_column(1, 1, 70)
        wsG.set_column(2, 2, 15)
        wsG.set_column(3, 3, 40)


        '''----------GEOMETRY CHECK----------'''
        #path for the second excel (checkGeometry)
        outTable = os.path.join(gdbPath, os.path.splitext(filename)[0])
        print("Running the check geometry tool on {} feature classes".format(len(filename_path)))
        checkResult = arcpy.CheckGeometry_management(filename_path, outTable)
        print("{} geometry problems found, see {} for details.".format(arcpy.GetCount_management(outTable)[0],
                                                                       outTable))

        #adding data to excel 2ndSheet (checKGeometry)
        checkRow = 1
        cursorTable = arcpy.da.SearchCursor(outTable, ["OBJECTID", "CLASS", "FEATURE_ID", "PROBLEM"])
        for row in cursorTable:
            #getting the shapeName to column
            wsG.write(checkRow, 1, filename, style)
            #translating to german
            if row[3] in d_problems:
                wsG.write(checkRow, 3, d_problems[row[3]], style)
            else:
                wsG.write(checkRow, 3, row[3], pink)
            wsG.write(checkRow, 0, row[0], style)
            wsG.write(checkRow, 2, row[2], style)
            checkRow += 1


        # setting the row variable for shapeName
        # iterate row in excel
        '''rowVariables'''
        rowField = 1
        rowPossible = 1
        '''rowVariables'''

        #searchCursor - comparing the values
        d_fieldsBestand = {}
        for i in fieldList_brut:
            if str(i.name.encode('latin-1')).upper()in d_req:
                d_fieldsBestand[str(i.name).upper()] = [[str(i.type), int(i.length)]]
        fieldList = [str((i.name).encode('latin-1').upper()) for i in arcpy.ListFields(filename_path) if str(i.name.encode('latin-1')).upper() in d_req]

        for i in range(0, len(fieldList)):
            d_fieldsBestand[fieldList[i]].append({})
        s = 0
        try:
            cursor = arcpy.da.SearchCursor(filename_path, fieldList)
            for row in cursor:
                for i in range(0, len(fieldList)):
                    if d_fieldsBestand[fieldList[i]][0][0] == 'Date':
                        valoare = str(row[i].day).rjust(2, "0") + "." + str(row[i].month).rjust(2, "0") + "." + str(
                            row[i].year)
                        if valoare not in d_fieldsBestand[fieldList[i]][1]:
                            d_fieldsBestand[fieldList[i]][1][valoare] = 1
                        else:
                            d_fieldsBestand[fieldList[i]][1][valoare] = d_fieldsBestand[fieldList[i]][1][valoare] + 1
                    elif d_fieldsBestand[fieldList[i]][0][0] == 'String':
                        try:
                            valoare = str(row[i])
                        except ValueError:
                            valoare = row[i]
                        if valoare not in d_fieldsBestand[fieldList[i]][1]:
                            d_fieldsBestand[fieldList[i]][1][valoare] = 1
                        else:
                            d_fieldsBestand[fieldList[i]][1][valoare] = d_fieldsBestand[fieldList[i]][1][valoare] + 1
                    else:
                        if row[i] not in d_fieldsBestand[fieldList[i]][1]:
                            d_fieldsBestand[fieldList[i]][1][row[i]] = 1
                        else:
                            d_fieldsBestand[fieldList[i]][1][row[i]] = d_fieldsBestand[fieldList[i]][1][row[i]] + 1

            del row, cursor
        except:
            print "searchCursor 1: Empty field list"
            #d_req["O_Shape"] = ["String", 100, str(filename)]
            #d_req["O_Shape"] = ["String", 100]

        #creating a dictionary for upcoming errors
        dictError = {}
        for i in d_req:
            dictError[i] = {}

        # validation cursor for the BL and Abschnitt fields
        try:
            cursor = arcpy.da.SearchCursor(filename_path, ["BL", "ABSCHNITT" ]) #"Abschnitt"
            for row in cursor:
                if row[0] == "BY" and row[1] in ["B", "C", "D", "AB", "CD", "G"]:
                    checking = 0
                elif row[0] == "DE" and row[1] in ["A", "B", "C", "D", "AB", "CD", "G"]:
                    checking = 0
                elif row[0] == "TH" and row[1] in ["A", "B", "C", "AB", "CD", "G"]:
                    checking = 0
                elif row[0] == "SN" and row[1] in ["A", "B", "C", "AB", "CD", "G"]:
                    checking = 0
                elif row[0] == "ST" and row[1] in ["A", "B", "AB", "G"]:
                    checking = 0
                else:
                    checking = 1
                    if row[0] in d_req["BL"][2:]:
                        checkError(dictError,"BL", "Abschnitt wert ungültiger:",row[1])
                    else:
                        checkError(dictError, "ABSCHNITT", "BL wert ungültiger:", row[0])
            del row, cursor
        except RuntimeError:
            print "searchCursor 2 (conditions of BL and Abs): Field not present in featureClass"

        # row var
        # adding data to xls columns
        for fieldReq in d_req:
            rowAttribute = rowField
            '''#########################  ADDING THE VALUES TO EXCEL    #############################################'''
            # adding the fieldsName for each shape
            if fieldReq == "BL":
                ws.write(rowField, nameCell, fieldReq, generalStyle)
            else:
                ws.write(rowField, nameCell, fieldReq.title(), generalStyle)
            # adding the type ofEachShape (datum or text)
            typeField = d_req[fieldReq][0]
            if typeField == "String":
                resType = "Text"
                lengthField = d_req[fieldReq][1]
            elif typeField == "Date":
                resType = "Datum"
                lengthField = ""
            ws.write(rowField, typeCell, resType + " : " + str(lengthField), generalStyle)
            # adding possible entries
            possibleEntry = d_req[fieldReq][2:]
            checking = 0

            '''------------------------------------------------------------------------------------------------------'''

            #writing in the dictError the "Existance of the field" error
            if fieldReq not in d_fieldsBestand:
                checkError(dictError,fieldReq, "Feld existiert nicht", None)
                checking = 1
            else:
                # check if is NULL or NOT
                if fieldReq not in ["U_SHAPE", "BEMERKUN", "ZONE", "EU_NR", "KATEGORIE", "ZONE", "GEBIET"]:
                    for nulitate in ["", " ", None]:
                        if nulitate in d_fieldsBestand[fieldReq][1]:
                            checkError(dictError, fieldReq, "Feld darf nicht leer sein:", None)
                            checking = 1

                if d_req[fieldReq][0] <> d_fieldsBestand[fieldReq][0][0]:  # Checking the Field TYPE
                    checkError(dictError, fieldReq, "Falscher Datentyp", typeField)
                    checking = 1
                else:
                    print (rowField, 7, "{}: Field Type is valid".format(fieldReq))
                if d_req[fieldReq][0] == "String":
                    if d_req[fieldReq][1] <> d_fieldsBestand[fieldReq][0][1]:  # Checking the Field LENGTH
                        print("{}: Field Length is NOT valid".format(fieldReq))
                        checkError(dictError, fieldReq, "Falsche Feldlänge:", (d_fieldsBestand[fieldReq][0][1]))
                        checking = 1
                    else:
                        print("{}: Field Length is valid".format(fieldReq))

                #iterating thorugh attributeValues of each field
                for attributeValue in d_fieldsBestand[fieldReq][1]:
                    #validation for "EU_Nr" field
                    if fieldReq == "EU_NR":
                        if attributeValue == " ":
                            checking = 0
                        elif len(attributeValue) == 8:
                            # defining the substrings
                            firsTwo = attributeValue[0:2]  # DE
                            spaceLine = attributeValue[2:3]  # " "
                            lastFive = attributeValue[3:8]  # 323
                            if ((firsTwo == "DE") and (spaceLine == " ") and int(lastFive)):
                                print "Attribute ------------------------is valid"
                                checking = 0;
                            else:
                                checkError(dictError, fieldReq, "Ungültiges attribut:",(str(attributeValue)))
                                print "Attribute -----------------------------is not valid"
                                checking = 1;
                        else:
                            checking = 1;
                            checkError(dictError, fieldReq, "Ungültiges attribut:", (str(attributeValue)))

                    if len(d_req[fieldReq]) > 2 and fieldReq <> "O_Shape":
                        try:
                            if attributeValue.encode('latin-1') not in d_req[fieldReq][2:]:
                                checking = 1
                                checkError(dictError, fieldReq, "Ungültiges attribut:", (str(attributeValue)))
                            else:
                                print "    ", attributeValue, "att: is valid #"
                            print("{} is valid\n".format(attributeValue.encode('latin-1')))
                        except ValueError:
                            print attributeValue, "is unicode value and could not be used into operation"

                # adding the attributes and the counts for the AttributeField
                try:
                    for i in (d_fieldsBestand[fieldReq][1]):
                        ws.write(rowAttribute, attCell, str(i.encode('latin-1')) + "  :  " + str(d_fieldsBestand[fieldReq][1][i]),style)
                        rowAttribute += 1
                except:
                    rowAttribute += 1


            #adding the "OK" or "NOT OK" Field
            #validaton of the fields (IF IT HAS PROBLEMS OR NOT)
            #adding the erros to the excel with the values as a list
            valueListError=[]
            if len(dictError[fieldReq]) > 0:
                for typeError in dictError[fieldReq]:
                    if dictError[fieldReq][typeError]==[None]:
                        valueListError.append(str(typeError).replace(":", ""))
                    else:
                        valueListError.append(str(typeError) + str(dictError[fieldReq][typeError]))
                ws.write(rowPossible, remarksCell, " | ".join(valueListError), pink) #
                ws.write(rowField, fieldCell, "NOT OK", pink) #
            else:
                ws.write(rowField, fieldCell, "OK", green) #

            #adding the Beschreibung column to excel
            for d in d_best:
                if d == fieldReq:
                    #print d
                    ws.write(rowField, beschreibungCell, d_best[d], style)

            # #validaton of the fields (IF IT HAS PROBLEMS OR NOT)
            # if checking <> 1:
            #     ws.write(rowField, fieldCell, "OK", green)
            # else:
            #     ws.write(rowField, fieldCell, "NOT OK",pink)

            rowField += 1

            # add the possibleValues to xls and align with the rest of the fields (rowField)
            for k in possibleEntry:
                ws.write(rowPossible, possibleCell, str(k + "\n"))
                rowPossible += 1
            rowField = max([rowField, rowPossible, rowAttribute])
            rowPossible = rowField
            #save the xlsx
            print("\n")
        wb.close()

        '''----------------------------------------------HEADER SHEET------------------------------------------------'''
        ##copy the spredsheet
        xl = Dispatch("Excel.Application")
        #xl.Visible = False  # You can remove this line if you don't want the Excel application to be visible
        wb1 = xl.Workbooks.Open(Filename=xlsPath)
        wb2 = xl.Workbooks.Open(Filename=tempxlsPath)
        ws2 = wb2.Worksheets(1)
        # primu se copie in woorkbookul din paranteze si specificat unde
        ws2.Copy(Before=wb1.Worksheets(1))
        wb1.CheckCompatibility = False
        wb1.DoNotPromptForConvert = True
        wb1.RefreshAll()
        attr = wb1.Worksheets(2)
        wsReport = wb1.Worksheets(1)
        wsReport.Cells(13,2).Value = filename
        if (code == 4647):
            wsReport.Cells(17,2).interior.color = rgb_to_hex((198, 239, 205))
            wsReport.Cells(17,2).Value = 4647
        else:
            wsReport.Cells(17,2).interior.color = rgb_to_hex((206, 198, 255))
            wsReport.Cells(17,2).Value = code
        wb1.Close(SaveChanges=True)
        xl.Quit()



print "----------------------------------------------------------------------------------------------------------------"
print "----------------------------------------------S C R I P T   F I N I S H E D-------------------------------------"
print "----------------------------------------------------------------------------------------------------------------"
