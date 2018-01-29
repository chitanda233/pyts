import xlwings as xw

def printName(name,wb):
    print( wb.sheets[name].name)
    print(wb.sheets[name].index)


app=xw.App(visible=False,add_book=False)
wb=app.books.add()
print(wb.sheets["sheet1"].index)

print(wb.sheets.count)
print(wb.sheets[wb.sheets.count-1].name)
wb.sheets.add(after=wb.sheets[wb.sheets.count-1])
#print(wb.sheets["toto"].index)

sht=wb.sheets.add(after=wb.sheets[wb.sheets.count-1])
sht.range('A1').value="第一个单元格"
wb.save('你好.xlsx')
wb.close()
app.quit()







