from ast import Not
import openpyxl
import pdfrw


excelFileName = "sellers.xlsx"
excelFileWorksheetName = "Form Responses 1"
templatePDF = "template.pdf"
completedPDF = "completed.pdf"
counter = 1

pdfValueMap = {
    "(wissell1)": "Wisconsin Seller Permit Number (15 digits starting with 456)",
    "(wissell1a)": "Wisconsin Seller Permit Number (15 digits starting with 456)",
    "(last4ssn1)": "SSN (last 4 digits)",
    "(last4fein1)": "FEIN (last 4 digits)",
    "(exempt1)": "Exemption code only if you are tax exempt",
    "(lglbusname1)": "Legal Business Name (if not sole proprietor)",
    "(vendcttnamlst1)": "Last Name",
    "(vendctnamfir1)": "First Name",
    "(vpn1)": "Phone Number",
    "(mailadd1)": "Mailing Address",
    "(emailadd1)": "Email Address",
    "(city1)": "City",
    "(state1)": "State",
    "(zip1)": "Zip",
    "(multi1)": "If a Multi-level Marketing company, please list the company name here (for number 2 responses)"
}

#----------------------------------------------------------------------------------
# Function to increment the counter and update keys
def increment_counter():
    global counter
    counter += 1
    # Reset counter to 1 when it reaches 4
    if counter > 4:
        counter = 1

#---------------------------------------------
# Function to create a new dictionary with updated keys
def incrementDict(excelRowDict):
    updated_dict = {}
    for key, value in excelRowDict.items():
        updated_key = key.replace("1", str(counter))
        updated_dict[updated_key] = value
    return updated_dict
#------------------

def parseExcelSheet():
    # Open the Excel file and select worksheet
    workbook = openpyxl.load_workbook(excelFileName)
    worksheet = workbook['Form Responses 1']

    # Get the headers from the first row (assuming headers are in the first row)
    headers = [cell.value for cell in worksheet[1] if cell.value is not None]

    # Initialize an empty list to store the excel rows
    ExcelRowsMappedToPDF = []

    # Loop through rows, starting from the second row (assuming the first row contains headers)
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        excelRowDict = {}
        for header, excelValue in zip(headers, row):
            for key, value in pdfValueMap.items():
                if value == header:
                    if value == "Wisconsin Seller Permit Number (15 digits starting with 456)":
                        if excelValue is None:
                            excelRowDict[key] = None
                        else:  
                            splitPermit = str(excelValue).split("-")
                            if len(splitPermit) > 2:
                                excelRowDict["(wissell1)"] = splitPermit[1]
                                excelRowDict["(wissell1a)"] = splitPermit[2]
                            else:
                                excelRowDict["(wissell1)"] = splitPermit[0]
                                excelRowDict["(wissell1a)"] = splitPermit[0]
                        #makes wissel1 the first value
                        # if len(splitPermit) > 0:
                        #     excelRowDict[key] = splitPermit[0]
                        # else:
                        #     excelRowDict[key] = None
                        break
                    else:
                        excelRowDict[key] = excelValue
                        break
            
        if excelRowDict["(emailadd1)"] is not None:
            incrementedDict = incrementDict(excelRowDict)
            ExcelRowsMappedToPDF.append(incrementedDict)
            increment_counter()
            #ExcelRowsMappedToPDF.append(excelRowDict)
            #replace num with counter

    # Close the Excel file when you're done
    workbook.close()

    return ExcelRowsMappedToPDF

#------------------------------------------------------------------------

def fillPDF(input_pdf_path, output_pdf_path, valueMapRows, start_page=2):
    template_pdf = pdfrw.PdfReader(input_pdf_path)
    counter = 0
    for page_number, page in enumerate(template_pdf.pages, 1):
        if page_number >= start_page:
            if "/Annots" in page and page.Annots:
                for annotation in page.Annots:
                    if annotation.get("/Subtype") == "/Widget" and annotation.get("/FT") == "/Tx":
                        field_name = str(annotation.get("/T"))
                        if not "pg" in field_name and not "dba" in field_name:
                            for row in valueMapRows:
                                if field_name in row:
                                    valueToUpdate = str(row[field_name])
                                    print(f"page number: {page_number} / annotation being updated: {field_name} with this value: {valueToUpdate} SUCCESS")
                                    annotation.update(pdfrw.PdfDict(V='{}'.format(valueToUpdate)))
                                    if "multi" in field_name:
                                        valueMapRows.remove(row)
                                    break
                                else:
                                    continue

        print('end of page')
    pdfrw.PdfWriter().write(output_pdf_path, template_pdf)


def fill_pdf_formz(input_pdf_path, output_pdf_path, data_dict):
    template_pdf = pdfrw.PdfReader(input_pdf_path)
    
    for page in template_pdf.pages:
        annotations = page.Annots
        if annotations is not None:
            for annotation in annotations:
                if annotation.get('/FT') == '/Tx':
                    field_name = annotation.get('/T')
                    print(field_name)
                    annotation.update(pdfrw.PdfDict(V='{}'.format(field_name)))
                    for key in data_dict:
                        if field_name in key:
                            annotation.update(pdfrw.PdfDict(V='{}'.format(key[field_name])))
                    if field_name in data_dict:
                        annotation.update(pdfrw.PdfDict(V='{}'.format(data_dict[field_name])))

    pdfrw.PdfWriter().write(output_pdf_path, template_pdf)

ExcelRowsMappedToPDF = parseExcelSheet()
# Access a mapped value using the key
# mapped_value = pdfValueMap["(pg2a)"]
# print(mapped_value)  # This will print "new_pg2a_value"
fillPDF(templatePDF, completedPDF, ExcelRowsMappedToPDF, start_page=2)
