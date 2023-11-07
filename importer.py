import openpyxl
import pdfrw

# Create a dictionary to map the values to other values
value_mapping = {
    "(pg2a)": "new_pg2a_value",
    "(pg2b)": "new_pg2b_value",
    "(wissell1)": "new_wissell1_value",
    "(wissell1a)": "new_wissell1a_value",
    "(last4ssn1)": "new_last4ssn1_value",
    "(last4fein1)": "new_last4fein1_value",
    "(exempt1)": "new_exempt1_value",
    "(lglbusname1)": "new_lglbusname1_value",
    "(dba1)": "new_dba1_value",
    "(vendcttnamlst1)": "new_vendcttnamlst1_value",
    "(vendctnamfir1)": "new_vendctnamfir1_value",
    "(vpn1)": "new_vpn1_value",
    "(mailadd1)": "new_mailadd1_value",
    "(emailadd1)": "new_emailadd1_value",
    "(city1)": "new_city1_value",
    "(state1)": "new_state1_value",
    "(zip1)": "new_zip1_value",
    "(multi1)": "new_multi1_value",
    "(wissell2)": "new_wissell2_value",
    "(wissell2a)": "new_wissell2a_value",
    "(last4ssn2)": "new_last4ssn2_value",
    "(last4fein2)": "new_last4fein2_value",
    "(exempt2)": "new_exempt2_value",
    "(lglbusname2)": "new_lglbusname2_value",
    "(dba2)": "new_dba2_value",
    "(vendcttnamlst2)": "new_vendcttnamlst2_value",
    "(vendctnamfir2)": "new_vendctnamfir2_value",
    "(vpn2)": "new_vpn2_value",
    "(mailadd2)": "new_mailadd2_value",
    "(emailadd2)": "new_emailadd2_value",
    "(city2)": "new_city2_value",
    "(state2)": "new_state2_value",
    "(zip2)": "new_zip2_value",
    "(multi2)": "new_multi2_value",
    "(wissell3)": "new_wissell3_value",
    "(wissell3a)": "new_wissell3a_value",
    "(last4ssn3)": "new_last4ssn3_value",
    "(last4fein3)": "new_last4fein3_value",
    "(exempt3)": "new_exempt3_value",
    "(lglbusname3)": "new_lglbusname3_value",
    "(dba3)": "new_dba3_value",
    "(vendcttnamlst3)": "new_vendcttnamlst3_value",
    "(vendctnamfir3)": "new_vendctnamfir3_value",
    "(vpn3)": "new_vpn3_value",
    "(mailadd3)": "new_mailadd3_value",
    "(emailadd3)": "new_emailadd3_value",
    "(city3)": "new_city3_value",
    "(state3)": "new_state3_value",
    "(zip3)": "new_zip3_value",
    "(multi3)": "new_multi3_value",
    "(wissell4)": "new_wissell4_value",
    "(wissell4a)": "new_wissell4a_value",
    "(last4ssn4)": "new_last4ssn4_value",
    "(last4fein4)": "new_last4fein4_value",
    "(exempt4)": "new_exempt4_value",
    "(lglbusname4)": "new_lglbusname4_value",
    "(dba4)": "new_dba4_value",
    "(vendcttnamlst4)": "new_vendcttnamlst4_value",
    "(vendctnamfir4)": "new_vendctnamfir4_value",
    "(vpn4)": "new_vpn4_value",
    "(mailadd4)": "new_mailadd4_value",
    "(emailadd4)": "new_emailadd4_value",
    "(city4)": "new_city4_value",
    "(state4)": "new_state4_value",
    "(zip4)": "new_zip4_value",
    "(multi4)": "new_multi4_value",
}

# Access a mapped value using the key
mapped_value = value_mapping["(pg2a)"]
print(mapped_value)  # This will print "new_pg2a_value"

#----------------------------------------------------------------------------------

# Open the Excel file
workbook = openpyxl.load_workbook('excel_doc.xlsx')  # TODO Replace 'excel_doc' with your excel file name

# Select the desired worksheet
worksheet = workbook['Sheet1']  # TODO Replace 'Sheet1' with your sheet name

# Get the headers from the first row (assuming headers are in the first row)
headers = [cell.value for cell in worksheet[1]]

# Initialize an empty list to store dictionaries
data_list = []

# Loop through rows, starting from the second row (assuming the first row contains headers)
for row in worksheet.iter_rows(min_row=2, values_only=True):
    row_dict = {}
    for header, value in zip(headers, row):
        if value_mapping:
            row_dict[header] = value
    data_list.append(row_dict)

# Close the Excel file when you're done
workbook.close()

# Print the list of dictionaries
for row_dict in data_list:
    print(row_dict)

#------------------------------------------------------------------------





def fill_pdf_form(input_pdf_path, output_pdf_path, data_dict):
    template_pdf = pdfrw.PdfReader(input_pdf_path)
    
    for page in template_pdf.pages:
        annotations = page.Annots
        if annotations is not None:
            for annotation in annotations:
                if annotation.get('/FT') == '/Tx':
                    field_name = annotation.get('/T')
                    print(field_name)
                    for key in data_dict:
                        if field_name in key:
                            annotation.update(pdfrw.PdfDict(V='{}'.format(key[field_name])))
                    # if field_name in data_dict:
                    #     annotation.update(pdfrw.PdfDict(V='{}'.format(data_dict[field_name])))

    pdfrw.PdfWriter().write(output_pdf_path, template_pdf)

# Example usage

# data_list = {
#     '(Name)': 'cody',
#     '(Address)': 'my address',
#     '(Phone Number)': '123-456-6798',
#     '(DOB)': '10/30/2023'
# }

fill_pdf_form('template.pdf', 'filled_output.pdf', data_list)
