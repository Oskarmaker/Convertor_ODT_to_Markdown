import docx

path = 'C:/Users/vovat/Desktop/АКМС/Практика 7.docx' # path to word file
doc = docx.Document(path)
doc.styles['Heading 1'].paragraph_format
print(type(doc))