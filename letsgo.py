import xlwings as xw

print("hello world")

app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False

wb = app.books.add()
<<<<<<< HEAD
=======

wb.sheets.add(after=wb.sheets[wb.sheets.count-1])
>>>>>>> b0c712d73939c76fb00af7b673fecb18c7621f52
wb.sheets.add(after=wb.sheets[wb.sheets.count-1])
print(wb.sheets.count)

wb.sheets[0].range('A1').value = "hello world"
wb.sheets[1].range('A1').value = "hello world"
<<<<<<< HEAD

wb.save(r"test.xlsx")
# wb2.save('test2.xlsx')
=======
>>>>>>> b0c712d73939c76fb00af7b673fecb18c7621f52

wb.save('test.xlsx')
wb.close()
app.quit()
print("goodbye world")
