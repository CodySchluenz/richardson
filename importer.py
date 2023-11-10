import openpyxl
import pdfrw


excelFileName = "sellers.xlsx"
excelFileWorksheetName = "Form Responses 1"
# Create a dictionary to map the values to other values
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
    "(state1)": "new_state1_value",
    "(zip1)": "State",
    "(multi1)": "Zip"
}

#----------------------------------------------------------------------------------

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
                    excelRowDict[key] = excelValue
                    break
            
            # if header in pdfValueMap.values():
            #     pdfValueMap[]
            #     print(pdfValueMap[str(header)])
            #     excelRowDict[header] = value
        if excelRowDict["(emailadd1)"] is not None:
            ExcelRowsMappedToPDF.append(excelRowDict)

    # Close the Excel file when you're done
    workbook.close()

    # Print the list of dictionaries
    # for excelRow in ExcelRowsMappedToPDF:
    #     print(excelRow)

#------------------------------------------------------------------------

column_dict = {
    'Timestamp': None,
    'Email Address': None,
    'Legal Business Name (if not sole proprietor)': None,
    'Last Name': None,
    'First Name': None,
    'Phone Number': None,
    'Mailing Address': None,
    'City': None,
    'State': None,
    'Zip': None,
    'Wisconsin Seller Permit Number (15 digits starting with 456)': None,
    'SSN (last 4 digits)': None,
    'FEIN (last 4 digits)': None,
    'Exemption code only if you are tax exempt': None,
    # Add more keys as needed
}

# Create a dictionary to map the values to other values
# value_mapping = {
#     "(pg2a)": "new_pg2a_value",
#     "(pg2b)": "new_pg2b_value",
#     "(wissell1)": data_list["Wisconsin Seller Permit Number (15 digits starting with 456)"].split("-")[1:1],  #
#     "(wissell1a)": data_list["Wisconsin Seller Permit Number (15 digits starting with 456)"].split("-")[2:],
#     "(last4ssn1)": data_list["SSN (last 4 digits)"],
#     "(last4fein1)": data_list["FEIN (last 4 digits)"],
#     "(exempt1)": data_list["Exemption code only if you are tax exempt"],
#     "(lglbusname1)": data_list["Legal Business Name (if not sole proprietor)"],
#     "(dba1)": data_list["No Collected, leave blank"],
#     "(vendcttnamlst1)": data_list["Last Name"],
#     "(vendctnamfir1)": data_list["First Name"],
#     "(vpn1)": data_list["Phone Number"],
#     "(mailadd1)": data_list["Mailing Address"],
#     "(emailadd1)": data_list["Email Address"],
#     "(city1)": data_list["City"],
#     "(state1)": data_list["State"],
#     "(zip1)": data_list["Zip"],
#     "(multi1)": data_list["If a Multi-level Marketing company, please list the company name here (for number 2 responses)"],

#     "(wissell2)": "new_wissell2_value",
#     "(wissell2a)": "new_wissell2a_value",
#     "(last4ssn2)": "new_last4ssn2_value",
#     "(last4fein2)": "new_last4fein2_value",
#     "(exempt2)": "new_exempt2_value",
#     "(lglbusname2)": "new_lglbusname2_value",
#     "(dba2)": "new_dba2_value",
#     "(vendcttnamlst2)": "new_vendcttnamlst2_value",
#     "(vendctnamfir2)": "new_vendctnamfir2_value",
#     "(vpn2)": "new_vpn2_value",
#     "(mailadd2)": "new_mailadd2_value",
#     "(emailadd2)": "new_emailadd2_value",
#     "(city2)": "new_city2_value",
#     "(state2)": "new_state2_value",
#     "(zip2)": "new_zip2_value",
#     "(multi2)": "new_multi2_value",
#     "(wissell3)": "new_wissell3_value",
#     "(wissell3a)": "new_wissell3a_value",
#     "(last4ssn3)": "new_last4ssn3_value",
#     "(last4fein3)": "new_last4fein3_value",
#     "(exempt3)": "new_exempt3_value",
#     "(lglbusname3)": "new_lglbusname3_value",
#     "(dba3)": "new_dba3_value",
#     "(vendcttnamlst3)": "new_vendcttnamlst3_value",
#     "(vendctnamfir3)": "new_vendctnamfir3_value",
#     "(vpn3)": "new_vpn3_value",
#     "(mailadd3)": "new_mailadd3_value",
#     "(emailadd3)": "new_emailadd3_value",
#     "(city3)": "new_city3_value",
#     "(state3)": "new_state3_value",
#     "(zip3)": "new_zip3_value",
#     "(multi3)": "new_multi3_value",
#     "(wissell4)": "new_wissell4_value",
#     "(wissell4a)": "new_wissell4a_value",
#     "(last4ssn4)": "new_last4ssn4_value",
#     "(last4fein4)": "new_last4fein4_value",
#     "(exempt4)": "new_exempt4_value",
#     "(lglbusname4)": "new_lglbusname4_value",
#     "(dba4)": "new_dba4_value",
#     "(vendcttnamlst4)": "new_vendcttnamlst4_value",
#     "(vendctnamfir4)": "new_vendctnamfir4_value",
#     "(vpn4)": "new_vpn4_value",
#     "(mailadd4)": "new_mailadd4_value",
#     "(emailadd4)": "new_emailadd4_value",
#     "(city4)": "new_city4_value",
#     "(state4)": "new_state4_value",
#     "(zip4)": "new_zip4_value",
#     "(multi4)": "new_multi4_value",
# }








def fillPDF(input_pdf_path, output_pdf_path, data_dict, start_page=2):
    template_pdf = pdfrw.PdfReader(input_pdf_path)

    for page_number, page in enumerate(template_pdf.pages, 1):
        if page_number >= start_page:
            if "/Annots" in page and page.Annots:
                for annotation in page.Annots:
                    if annotation.get("/Subtype") == "/Widget" and annotation.get("/FT") == "/Tx":
                        field_name = annotation.get("/T")
                        print(field_name)
                        annotation.update(pdfrw.PdfDict(V='{}'.format(field_name)))
                        #if field_name in data_dict:
                            #annotation.update(pdfrw.PdfDict(V='{}'.format(data_dict[field_name])))

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

# Example usage

# data_list = {
#     '(Name)': 'cody',
#     '(Address)': 'my address',
#     '(Phone Number)': '123-456-6798',
#     '(DOB)': '10/30/2023'
# }

parseExcelSheet()
# Access a mapped value using the key
mapped_value = pdfValueMap["(pg2a)"]
print(mapped_value)  # This will print "new_pg2a_value"
fillPDF('template.pdf', 'filled_output.pdf', data_list, start_page=2)
