# The employees data will be in excel format:
All employee information should be on one raw
![excel data iniput](https://github.com/MoSbeaa/Excel-to-PDF-Form-Filler_Py/blob/main/src/excel%20data%20iniput.png)
<img src="https://github.com/MoSbeaa/Excel-to-PDF-Form-Filler_Py/blob/main/src/excel%20data%20iniput.pn" width="200" />
# The pdf form template:
Fields should be created using Adobe DC, and the fileds names should match the excel columns names, see below for more details
![pdf form input](https://github.com/MoSbeaa/Excel-to-PDF-Form-Filler_Py/blob/main/src/pdf%20form%20input.png)

# The files in the output folder
Each raw from excel file generate one pdf file
![Output files](https://github.com/MoSbeaa/Excel-to-PDF-Form-Filler_Py/blob/main/src/Output%20files.png)

# The file after filling will be look like
![pdf form output](https://github.com/MoSbeaa/Excel-to-PDF-Form-Filler_Py/blob/main/src/pdf%20form%20output.png)



# How to Start?

## 1 Change the variables names
  variables that need to change when you have new data report is:
  * rawDataFile_filename  --> this variable holds the input excel file of row data, and this file in directory should be in subfolder in the main code folder --> usually is in the "In" subfolder.
  * pdf_template  --> this variable holds the pdf template file, and this file in directory should be in subfolder in the main code folder --> usually is in the "In" subfolder.
  * pdf_outPut_path  --> this variable holds the output folder --> usually is "Out" subfolder.
  * outPutFileName  --> this variable holds the output file name, change it as you like

## 2 Check the Dictionary variable
  * dictionary --> this dictionary holds all variables name from the pdf form and the value from excel file, it will be filled automatically, but
  * If you have field name that is not same as columns name in excel file, you should to add it manually, otherwise the code will not work.
  * see dictionary value to know how to add manually.

## 3 Make the fields in the PDF
  
  * You should use tool called 'Prepare Form' in adobe acrobat pro DC.
  * Important: before start adding fields, make sure to delete all auto created fields in the form, you can look on the field 
  panel at the right to see all the fields existed.
  * Use 'Add a text filed' option, not 'Add text', Then assign the fields places.
  * Be sure that the fields name is exact as the column name in excel file
  * Also, any font format can be made during making the field.

| :exclamation:  This is very important   |
|-----------------------------------------|
  > No field name duplication allowed: if you duplicated any filed name it will not be filled (at least in my code)
