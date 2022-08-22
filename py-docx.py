import docx

h = r'C:\Users\Иван\Desktop\Иван\python\pythonProject\Word\жждьждь.docx'
doc = docx.Document(h)
par = doc.paragraphs

print(len(par))
print(par)

for t in range(len(par[0].runs)):
    print('l', par[0].runs[t].text)

print(dir(docx))

