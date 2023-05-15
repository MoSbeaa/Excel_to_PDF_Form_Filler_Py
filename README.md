![alt text](https://github.com/MoSbeaa/Excel-to-PDF-Form-Filler_Py/tree/main/src/excel data iniput.png?raw=true)



# How to Start?

# 1 Change the variables names
  variables that need to change when you have new data report is:

  > rawDataFile_filename  --> this variable holds the input excel file of row data, and this file in directory should be in subfolder in the main code folder --> usually is in the "In" subfolder.

  > pdf_template  --> this variable holds the pdf template file, and this file in directory should be in subfolder in the main code folder --> usually is in the "In" subfolder.

  > pdf_outPut_path  --> this variable holds the output folder --> usually is "Out" subfolder.

  > outPutFileName  --> this variable holds the output file name, change it as you like

# 2 Check the Dictionary variable
  > dictionary --> this dictionary holds all variables name from the pdf form and the value from excel file, it will be filled automatically, but
  > If you have field name that is not same as columns name in excel file, you should to add it manually, otherwise the code will not work.
  > see dictionary value to know how to add manually.

# 3 Make the fields in the PDF
  
  > You should use tool called 'Prepare Form' in adobe acrobat pro DC.
  > Important: before start adding fields, make sure to delete all auto created fields in the form, you can look on the field 
  panel at the right to see all the fields existed.
  > Use 'Add a text filed' option, not 'Add text', Then assign the fields places.
  > Be sure that the fields name is exact as the column name in excel file
  > Also, any font format can be made during making the field.

## Warning 
  > No field name duplication allowed: if you duplicated any filed name it will not be filled (at least in my code)
