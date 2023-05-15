
import os
import pdfrw
import pandas as pd
 
rowDataFile_filename = "./Employees_Information_Data.xlsx" # Excel file path
pdf_template = "./Employees_Information_Template.pdf" # PDF file path
pdf_outPut_path = './Out' # output folder path



def write_new_pdf(pdf_template, output_pdf_path, columnName):
    ANNOT_KEY = '/Annots'           # key for all annotations within a page
    ANNOT_FIELD_KEY = '/T'          # Name of field. i.e. given ID of field
    ANNOT_FORM_type = '/FT'         # Form type (e.g. text/button)
    ANNOT_FORM_button = '/Btn'      # ID for buttons, i.e. a checkbox
    ANNOT_FORM_text = '/Tx'         # ID for textbox
    SUBTYPE_KEY = '/Subtype'
    WIDGET_SUBTYPE_KEY = '/Widget'
    pdfin = os.path.normpath(os.path.join(os.getcwd(),'in',pdf_template)) # full directory of pdf output file
    template_pdf = pdfrw.PdfReader(open(pdfin, "rb"))  # make new pdf file and hold it

    dictionary = { 
        # in case the filed doesn't have column in the excel data, you can enter it manually
        # Pattern -->  "field Name": "any value you want"
        # in case you want to add it manually and cast value from excel 
        # you can call value from the excel itself use -->  "field Name": columnName["column name"]
        
        
        
            }
  
    ## the next for loop will fill the dictionary with all data required to fill the pdf form
    for Page in template_pdf.pages:
        if Page[ANNOT_KEY]:
            for annotation in Page[ANNOT_KEY]:
                if annotation[ANNOT_FIELD_KEY] and annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY :
                    key = annotation[ANNOT_FIELD_KEY][1:-1]
                    if key in dictionary:
                        continue
                    if annotation[ANNOT_FORM_type] == ANNOT_FORM_button:
                        # button field i.e. a checkbox
                        dictionary[key] = columnName[key]
                        annotation.update(pdfrw.PdfDict(V=pdfrw.PdfName(dictionary[key]) , AS=pdfrw.PdfName(dictionary[key]) ))
                    elif annotation[ANNOT_FORM_type] == ANNOT_FORM_text:
                        # regular text field
                        dictionary[key] = columnName[key]
                        annotation.update( pdfrw.PdfDict( V=dictionary[key], AP=dictionary[key]) )
                    
                    annotation.update(pdfrw.PdfDict(Ff=1)) # this line to make the file fields none editable
                    
    # this will fill the pdf form field with the data
    template_pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))

    # this line will release the final version of the filled pdf form to the output path
    pdfrw.PdfWriter().write(output_pdf_path, template_pdf)

def get_data_from_excel(pdf_template,pdf_outPut_path, data):
    i = 0 
    for j, columnName in data.iterrows(): #j is important, j is number of rows
        i += 1
        ## the output file name
        outPutFileName = columnName['Full_name'] + "_information"
        ######################
        temp_out_dir = os.path.normpath(os.path.join(pdf_outPut_path,outPutFileName +'.pdf'))
        write_new_pdf(pdf_template, temp_out_dir, columnName)
        print(i, "...please wait...")
    return i

# Code start here
try:
    rowDataFilein = os.path.normpath(os.path.join(os.getcwd(),'in',rowDataFile_filename)) # this line will get the full directory of rawdata in my machine
    data = pd.read_excel(rowDataFilein) # this will make stream holds all the rawdata values 
    rowDataFile_fields = data.columns.tolist() # this will make list of all columns header

    if __name__ == '__main__': 
        run = get_data_from_excel(pdf_template,pdf_outPut_path, data)
        print("Finish\n{} files has been created successfully".format(run))
            
except PermissionError as pr:
    print("Please close this file \" {} \" \nand try again".format(pr.filename))
except NameError as nerr:
    print()
except OSError as orr:
    print("Please check of the following directory \" {} \" \nmaybe it has typo error".format(orr.filename))
except FileNotFoundError as ferr:
    print("Problem with the following directory,\n{} \ncheck if this right directory for your rawdata and pdf template".format(ferr.filename))
except KeyError as msg :
    m1 = "The field name {} it has error, please check these scenarios\n".format(msg)
    m2 = "1- Probably it has typo error in pdf form\n"
    m3 = "2- If this field name is not same as the excel's column name, then you should add it manually in the \'dictionary\'"
    print(m1,m2,m3)
except AttributeError as atterr:
    print("You are using \"{}\" function incorrectly the error statement is\n{}".format(atterr.name,atterr.args))