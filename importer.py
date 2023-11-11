import openpyxl
import pdfrw

EXCEL_FILE_NAME = "sellers.xlsx"
EXCEL_WORKSHEET_NAME = "Form Responses 1"
TEMPLATE_PDF = "template.pdf"
COMPLETED_PDF = "completed.pdf"
VENDORS_PER_PAGE = 4
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
# Function to increment the counter to update the pdfValueMap key with a new number between 1-4
# This is needed because the pdf fields are in the range of 1-4. 
# for example, (wissell1), (wissell2), (wissell3), (wissell4).
def increment_counter():
    global counter
    counter += 1
    # Reset counter to 1 when it reaches the max number of vendors per page.
    if counter > VENDORS_PER_PAGE:
        counter = 1

#---------------------------------------------
# Function to create a new dictionary with updated keys from the increment_counter function.
def incrementDict(excelRowDict):
    updated_dict = {}
    for key, value in excelRowDict.items():
        updated_key = key.replace("1", str(counter))
        updated_dict[updated_key] = value
    return updated_dict
#------------------

def parseExcelSheet():
    # Open the Excel file and select worksheet
    workbook = openpyxl.load_workbook(EXCEL_FILE_NAME)
    worksheet = workbook[EXCEL_WORKSHEET_NAME]

    # Get the headers from the first row (assuming headers are in the first row)
    headers = [cell.value for cell in worksheet[1] if cell.value is not None]

    # Initialize an empty list to store the excel rows
    ExcelRowsMappedToPDF = []

    # Loop through rows, starting from the second row (assuming the first row contains headers)
    # builds a key value map between the excel rows and the PDF document fields.
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
                        break
                    else:
                        excelRowDict[key] = str(excelValue)
                        break
            
        if excelRowDict["(emailadd1)"] is not None:
            incrementedDict = incrementDict(excelRowDict)
            ExcelRowsMappedToPDF.append(incrementedDict)
            increment_counter()

    # Close the Excel file when you're done
    workbook.close()
    return ExcelRowsMappedToPDF

#------------------------------------------------------------------------
# Function to update the PDF with 
def updatePDF(input_pdf_path, output_pdf_path, valueMapRows, start_page=2):
    template_pdf = pdfrw.PdfReader(input_pdf_path)
    for page_number, page in enumerate(template_pdf.pages, 1):
        if page_number >= start_page:
            if "/Annots" in page and page.Annots:
                for annotation in page.Annots:
                    if annotation.get("/Subtype") == "/Widget" and annotation.get("/FT") == "/Tx":
                        field_name = str(annotation.get("/T"))
                        # if skip the pg and dba fields because we dont use them.
                        if not "pg" in field_name and not "dba" in field_name:
                            for row in valueMapRows:
                                if field_name in row:
                                    valueToUpdate = str(row[field_name])
                                    print(f"page number: {page_number} / annotation being updated: {field_name} with this value: {valueToUpdate} SUCCESS")
                                    annotation.update(pdfrw.PdfDict(V='{}'.format(valueToUpdate)))
                                    # multi is the last field per vendor so if the field_name is multi, we can say we are at the end of the vendor and 
                                    # need to remove the completed vendor to start the new vendor.
                                    if "multi" in field_name:
                                        valueMapRows.remove(row)
                                    break
    # writes the changes to the pdf
    pdfrw.PdfWriter().write(output_pdf_path, template_pdf)


if __name__ == "__main__":
    ExcelRowsMappedToPDF = parseExcelSheet()
    updatePDF(TEMPLATE_PDF, COMPLETED_PDF, ExcelRowsMappedToPDF, start_page=2)