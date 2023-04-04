from docx.api import Document
import pandas








from docx.api import Document
import openpyxl
# Load the first table from your document. In your example file,
# there is only one table, so I just grab the first one.
document = Document('D:/agsk10.docx')
table = document.tables[0]

# Data will be a list of rows represented as dictionaries
# containing each row's data.
data = []

keys = None
for i, row in enumerate(table.rows):
    
    text = (cell.text for cell in row.cells)
    print(text)
   
    if i == 0:
        keys = tuple(text)
        continue

   
    row_data = dict(zip( text))
    data.append(row_data)

df = pandas.DataFrame(data)
print('this is data from word file = ',data)
df.to_excel('D:/test10.xlsx', sheet_name='1')


# df = pandas.DataFrame([[11, 21, 31], [12, 22, 32], [31, 32, 33]],
#                   index=['one', 'two', 'three'], columns=['a', 'b', 'c'])
# print(df)
# df.to_excel('D:/test.xlsx', sheet_name='1')

# df.columns = ["Column1", "Column2","Column3", "Column4"]
# df.to_excel("D:/test.xlsx")

