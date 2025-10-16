import docx
from styls import Styles

path = 'C:/Users/vovat/Desktop/АКМС/Тест1.docx' # path to word file
doc = docx.Document(path)
styles = Styles(doc)
styles.run()
doc.save(path)