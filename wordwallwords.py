from PyDictionary import PyDictionary
import os
import sys
import docx
doc = docx.Document()
dictionary = PyDictionary()
word = input('Please type the words you want, seperated by a comma: ')
word = word.replace(" ", "")
fileName = ''
allwords = word.split(',')
for x in range(len(allwords)):
    print(allwords[x])
verify = input('Just making sure, are those the words you chose? (Type yes or no)')
fixeddefs = []
if verify == 'yes':
    fileName = input('What name do you want for the file? Make sure it is a unique name: ')
    for x in allwords:
        response = str(dictionary.meaning(x, disable_errors=True))
        definition = response.replace("{", '')
        definition = definition.replace("}", '')
        definition = definition.split("]", 1)
        definition =definition[0]
        definition = definition.replace("'", "")
        definition = definition.replace("[", "")
        definition = definition.replace("]", "")
        definition = definition.replace(":", "|")
      
        if definition == 'None':
            definition = 'Is|Not a valid Word'
        definition = x + '| ' + definition
        definitionarr = definition.split("|")
        fixeddefs.append(definitionarr)
elif verify == 'no':
    print("Please restart the program and try again")
menuTable = doc.add_table(rows=1, cols=3)
menuTable.style = 'Table Grid'

hdr_Cells = menuTable.rows[0].cells
hdr_Cells[0].text = 'Word'
hdr_Cells[1].text = 'POS'
hdr_Cells[2].text = 'Definition'
for word, pos, defin in fixeddefs:
    row_Cells = menuTable.add_row().cells
    row_Cells[0].text = word
    row_Cells[1].text = pos
    row_Cells[2].text = defin
doc.save(fileName+'.docx')
os.system("start " + fileName+".docx")
