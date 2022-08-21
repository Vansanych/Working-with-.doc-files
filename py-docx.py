import docx

h = r'C:\Users\Иван\Desktop\Иван\python\pythonProject\Word\жждьждь.docx'
doc = docx.Document(h)
par = doc.paragraphs

print(len(par))
print(par)

for t in range(len(par[0].runs)):
    print('l', par[0].runs[t].text)

print(dir(docx))


# print(docx.text(h))

# print(par[0].runs[0].text.Font.name)

# print(par[0].runs[0].text)
# print(par[0].runs[1].text)
# tab = doc.tables
# print(tab)
# print(dir(docx.Document))
#
# print(dir(docx.Document.__dict__))
# print(dir(docx.Document()))
# print(dir(docx.Document()) == dir(doc))
# print(dir(doc.tables))
#
# print(dir(doc.tables[0]))
# print(dir(doc.tables[0].rows[0]))
# print(dir(doc.tables[0].rows[0].cells[0]))
#
# print(dir(doc.tables[0].rows[0].cells[0].text))
# print('text', doc.tables[0].rows[1].cells[0].text)
# print('text', doc.tables[0].rows[1].cells[1].text)
# # doc.tables[0].rows[1].cells[1].text = 'hbh'
# print('text', doc.tables[0].rows[1].cells[1].text)
#
# doc.save(r'C:\Users\Иван\Desktop\Иван\python\pythonProject\Word\11.docx')
