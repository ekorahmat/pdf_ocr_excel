import os
from pdf2image import convert_from_path
import function_crop_image as ci
import function_get_text as gt
import function_excel as fe
from datetime import datetime
import xlsxwriter

# SET POPPLER PATH
# POPPLER IS A PDF RENDERING LIBRARY
# USED TO CONVERT PDF FILE TO IMAGE FILE IN THIS PROGRAM
# HOW TO INSTALL POPPLER IN YOUR OPERATING SYSTEM, PLEASE SEE THE DOCUMENTATION IN
# https://poppler.freedesktop.org/
# ADJUST THE PATH OF THE POPPLER LIBRARY ACCORDING WHERE ITS INSTALLED
popplerPath = r'C:\Program Files\poppler-0.68.0_x86\poppler-0.68.0\bin'
print('popplerPath = ', popplerPath)

# GET CURRENT ROOT DIRECTORY WHERE PROGRAM RUNS
currDir = os.getcwd()
print('currDir = ', currDir)

# GET - PDF FOLDER
# PDF FOLDER CONTAINS PDF FILES THAT WANT TO CONVERT TO EXCEL FILE
pdfFolder = currDir + "\\pdf"
print('pdfFolder = ', pdfFolder)

# GET - LIST OF PDF FILES IN PDF FOLDER
pdfFiles = os.listdir(pdfFolder)
print('pdfFiles = ', pdfFiles)

# GET CONVERT PDF TO IMAGE FOLDER
# TO CONVERT PDF FILE TO EXCEL FILE, WE NEED TO CONVERT PDF FILE TO IMAGE FILE
convertPdf2ImageFolder = currDir + "\\convert_pdf_2_image"
print('convertPdf2ImageFolder = ', convertPdf2ImageFolder)
if not os.path.exists(convertPdf2ImageFolder):
    os.makedirs(convertPdf2ImageFolder)

# CONVERT ALL PDF FILES TO IMAGE FILE AND SAVE IT TO CONVERT PDF 2 IMAGE FOLDER
for pdfFile in pdfFiles:
    # GET PDF FILE NAME
    print('pdfFile = ', pdfFile)
    # SET PDF PATH FILE NAME
    pdfFilePath = pdfFolder + "\\" + pdfFile
    print('pdfFilePath = ', pdfFilePath)
    # SET - CONVERTED IMAGE FILE
    pdfImageConvertedFile = pdfFile.replace("pdf", "jpg")
    print('pdfImageConvertedFile = ', pdfImageConvertedFile)
    # SET - PATH CONVERTED IMAGE FILE
    pdfImageConvertedFilePath = convertPdf2ImageFolder + "\\" + pdfImageConvertedFile
    print('pdfImageConvertedFilePath = ', pdfImageConvertedFilePath)

    # CONVERT PDF FILE TO IMAGE FILE
    try:
        convertPdf2Image = convert_from_path(pdfFilePath, 500, poppler_path=popplerPath)
        # CONVERT ONLY ONE PAGE FILE
        if(len(convertPdf2Image) == 1):
            # SAVE TO IMAGE FILE
            convertPdf2Image[0].save(pdfImageConvertedFilePath, 'JPEG')
    except Exception as e:
        print('e = ', e)

# GET LIST CONVERTED IMAGE FILES
convertedImageFiles = os.listdir(convertPdf2ImageFolder)
print('convertedImageFiles = ', convertedImageFiles)

# THE NEXT STEP AFTER CONVERT PDF FILE TO IMAGE FILE
# IS TO GET TEXT FROM IMAGE FILE TO BE SAVED AS DATA TO THE TABLE IN EXCEL FILE
# WE GET TEXT BY CROPPED THE IMAGES TO SEVERAL IMAGES

# GET CROPPED IMAGES FOLDER
croppedImgesFolder = currDir + "\\cropped_images"
print('croppedImgesFolder = ', croppedImgesFolder)
if not os.path.exists(croppedImgesFolder):
    os.makedirs(croppedImgesFolder)

# SET - TEXT TO CAPTURE FROM IMAGE FILE
textToCapture = []
textToCapture.append("PO DATE")
textToCapture.append("PO NUMBER")
textToCapture.append("COMPANY NAME")
textToCapture.append("COMPANY ADDRESS")
textToCapture.append("COMPANY CITY")
textToCapture.append("COMPANY PHONE NUMBER")
textToCapture.append("COMPANY WEB SITES")
textToCapture.append("VENDOR NAME")
textToCapture.append("VENDOR ADDRESS")
textToCapture.append("VENDOR CITY")
textToCapture.append("VENDOR PHONE NUMBER")
textToCapture.append("SHIP TO PERSON")
textToCapture.append("SHIP TO COMPANY NAME")
textToCapture.append("SHIP TO COMPANY ADDRESS")
textToCapture.append("SHIP TO COMPANY CITY")
textToCapture.append("SHIP TO COMPANY PHONE NUMBER")
textToCapture.append("DELIVERY DATE")
textToCapture.append("SUBTOTAL")
textToCapture.append("TAX")
textToCapture.append("SHIPPING")
textToCapture.append("TOTAL")

# DATA TO SAVED
savedData = []

# CROPP IMAGE TO SEVERAL IMAGES
# GET DATA FROM CROPPED IMAGES
for convertedImageFile in convertedImageFiles:
    print('convertedImageFile = ', convertedImageFile)

    convertedImageFilePath = convertPdf2ImageFolder + "\\" + convertedImageFile
    print('convertedImageFilePath = ', convertedImageFilePath)

    # SET DATA FROM TEXT
    data = []
    dataFile = ("CONVERTED IMAGE FILE", convertedImageFile)
    data.append(dataFile)

    # SET FILE NAME CROPPED IMAGE - OCR - TEXT
    for capture in textToCapture:
        cropp = capture
        capture = capture.lower()
        capture = capture.replace(" ", "_")
        # print('capture = ', capture)
        croppedImageFile = convertedImageFile.replace(".jpg","")
        croppedImageFile = croppedImageFile + "_" + capture + ".jpg"
        # print('croppedImageFile = ', croppedImageFile)
        croppedImageFilePath = croppedImgesFolder + "\\" + croppedImageFile
        print('croppedImageFilePath = ', croppedImageFilePath)

        # CROPP IMAGE
        resultCroppedImage = ci.croppImage(convertedImageFilePath, croppedImageFilePath, cropp)
        print('resultCroppedImage = ', resultCroppedImage)
        # GET TEXT FROM CROPPED IMAGE
        if (resultCroppedImage == "SUCCESS"):
            resultGetText = gt.getSimpleTextFromImage(croppedImageFilePath)
            print('resultGetText = ', resultGetText)
            if (resultGetText != ""):
                if(capture == "company_city"):
                    results = resultGetText.split(", ")
                    dataText = (capture, results[0])
                    data.append(dataText)
                    dataText = ("company_zip_code", results[2])
                    data.append(dataText)
                elif(capture == "vendor_city"):
                    results = resultGetText.split(", ")
                    dataText = (capture, results[0])
                    data.append(dataText)
                    dataText = ("vendor_zip_code", results[2])
                    data.append(dataText)
                elif(capture == "ship_to_company_city"):
                    results = resultGetText.split(", ")
                    dataText = (capture, results[0])
                    data.append(dataText)
                    dataText = ("ship_to_compnay_zip_code", results[2])
                    data.append(dataText)
                else:
                    dataText = (capture, resultGetText)
                    data.append(dataText)                
        else:
            exit()

    savedData.append(data)
    # DELETE ALL CONVERTED PDF 2 IMAGE FILE
    os.remove(convertedImageFilePath)
    print('delete File convertedImageFilePath = ', convertedImageFilePath)

# NEXT STEP IS TO GENERATE EXCEL FILE
# AND INSERT DATA INTO TABLE 

# GET CURRENT DATE TIME
currDateTime = datetime.now().strftime('%y%m%d_%H%M%S')
print(currDateTime)

# GENERATED EXCEL FILE
excelFile = "pdf_to_excel_" + currDateTime + ".xlsx"
print('excelFile = ', excelFile)

# GET EXCEL FOLDER
excelFolder = currDir + "\\excel"
if not os.path.exists(excelFolder):
    os.makedirs(excelFolder)

# GET EXCEL FILE PATH
excelFilePath = excelFolder + "\\" + excelFile
print('excelFilePath = ', excelFilePath)

# CREATE EXCEL FILE
workbook = xlsxwriter.Workbook(excelFilePath)
# SET EXCEL SHEET NAME
sheetName = "Purchase Order Data"
worksheet = workbook.add_worksheet(sheetName)

# MERGE FORMAT - FOR COLUMN HEADER TABLE
merge_format = workbook.add_format({
    'align' : 'center',
    'valign' : 'vcenter'
})

# SET HEADER COLUMN NUMBER
columnNumber = "Number"
# MERGE CELL A1:A2
worksheet.merge_range('A1:A2', "Number", merge_format)

# SET HEADER COLUMN PO Number
columnPONumber = "PO Number"
# MERGE CELL B1:B2
worksheet.merge_range('B1:B2', "PO Number", merge_format)
worksheet.set_column(1, 1, 22)

# SET HEADER COLUMN PO Date
columnPODate = "PO Date"
# MERGE CELL C1:C2
worksheet.merge_range('C1:C2', "PO Date", merge_format)
worksheet.set_column(2, 2, 12)

# SET HEADER COLUMN COMPANY
# MERGE CELL D1:I1
worksheet.merge_range('D1:I1', "Company", merge_format)

# SET HEADER COLUMN COMPANY - NAME
worksheet.write('D2', 'Name')
worksheet.set_column(3, 3, 12)

# SET HEADER COLUMN COMPANY - ADDRESS
worksheet.write('E2', 'Address')
worksheet.set_column(4, 4, 40)

# SET HEADER COLUMN COMPANY - CITY
worksheet.write('F2', 'City')
worksheet.set_column(5, 5, 12)

# SET HEADER COLUMN COMPANY - ZIP CODE
worksheet.write('G2', 'Zip Code')
worksheet.set_column(6, 6, 12)

# SET HEADER COLUMN COMPANY - PHONE NUMBER
worksheet.write('H2', 'Phone Number')
worksheet.set_column(7, 7, 15)

# SET HEADER COLUMN COMPANY - WEB SITES
worksheet.write('I2', 'Web Sites')
worksheet.set_column(8, 8, 15)

# SET HEADER COLUMN VENDOR
# MERGE CELL J1:N1
worksheet.merge_range('J1:N1', "Vendor", merge_format)

# SET HEADER COLUMN VENDOR - NAME
worksheet.write('J2', 'Name')
worksheet.set_column(9, 9, 12)

# SET HEADER COLUMN VENDOR - ADDRESS
worksheet.write('K2', 'Address')
worksheet.set_column(10, 10, 40)

# SET HEADER COLUMN VENDOR - CITY
worksheet.write('L2', 'City')
worksheet.set_column(11, 11, 12)

# SET HEADER COLUMN VENDOR - ZIP CODE
worksheet.write('M2', 'Zip Code')
worksheet.set_column(12, 12, 12)

# SET HEADER COLUMN VENDOR - PHONE NUMBER
worksheet.write('N2', 'Phone Number')
worksheet.set_column(13, 13, 15)

# SET HEADER COLUMN SHIP TO COMPANY
# MERGE CELL O1:T1
worksheet.merge_range('O1:T1', "Ship to Company", merge_format)

# SET HEADER COLUMN SHIP TO COMPANY - PERSON
worksheet.write('O2', 'Person')
worksheet.set_column(14, 14, 12)

# SET HEADER COLUMN SHIP TO COMPANY - NAME
worksheet.write('P2', 'Name')
worksheet.set_column(15, 15, 12)

# SET HEADER COLUMN SHIP TO COMPANY - ADDRESS
worksheet.write('Q2', 'Address')
worksheet.set_column(16, 16, 40)

# SET HEADER COLUMN SHIP TO COMPANY - CITY
worksheet.write('R2', 'City')
worksheet.set_column(17, 17, 12)

# SET HEADER COLUMN SHIP TO COMPANY - ZIP CODE
worksheet.write('S2', 'Zip Code')
worksheet.set_column(18, 18, 12)

# SET HEADER COLUMN SHIP TO COMPANY - PHONE NUMBER
worksheet.write('T2', 'Phone Number')
worksheet.set_column(19, 19, 15)

# SET HEADER COLUMN DELIVERY DATE
# MERGE CELL U1:U2
worksheet.merge_range('U1:U2', "Delivery Date", merge_format)
worksheet.set_column(20, 20, 12)

# SET HEADER COLUMN SUBTOTAL
# MERGE CELL V1:V2
worksheet.merge_range('V1:V2', "Subtotal", merge_format)
worksheet.set_column(21, 21, 15)

# SET HEADER COLUMN TAX
# MERGE CELL W1:W2
worksheet.merge_range('W1:W2', "Tax", merge_format)
worksheet.set_column(22, 22, 6)

# SET HEADER COLUMN SHIPPING
# MERGE CELL X1:X2
worksheet.merge_range('X1:X2', "Shipping", merge_format)
worksheet.set_column(23, 23, 15)

# SET HEADER COLUMN TOTAL
# MERGE CELL Y1:Y2
worksheet.merge_range('Y1:Y2', "Total", merge_format)
worksheet.set_column(24, 24, 15)

# START ROW
row = 2
# START NUMBER
number = 1

# GET ALL DATA AND INSERT THE DATA INTO CELL VALUE IN EXCEL FILE
for data in savedData:
    print('data = ', data)

    # WRITE NUMBER INTO TABLE
    worksheet.write(row, 0, number)

    # WRITE DATA INTO TABLE
    for dataDetail in data:
        print('dataDetail = ', dataDetail)
        cellColumn = fe.getCellColumn(dataDetail[0])
        print('cellColumn = ', cellColumn)
        if cellColumn >= 0:
            print('dataDetail[1] = ', dataDetail[1])
            if(dataDetail[0] == "subtotal" or dataDetail[0] == "shipping" or dataDetail[0] == "total"):
                value = dataDetail[1].replace("S", "")
                value = value.replace(",", "")
                value = float(value)
                worksheet.write(row, cellColumn, value)
            elif(dataDetail[0] == "po_number"):
                value = dataDetail[1].upper()
                worksheet.write(row, cellColumn, value)
            else:
                worksheet.write(row, cellColumn, dataDetail[1])

    row = row + 1
    number = number + 1

# DELETE ALL CROPPED IMAGE FILES
croppedImageFiles = os.listdir(croppedImgesFolder)
for fileToDelete in croppedImageFiles:
    print('fileToDelete = ', fileToDelete)
    fileToDeletePath = croppedImgesFolder + "\\" + fileToDelete
    os.remove(fileToDeletePath)

# CLOSE EXCEL FILE
workbook.close()
