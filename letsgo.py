import xlwings as xw

print("hello world")

app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False

wb = app.books.add()
wb2 = app.books.add()

wb.sheets.add("yahaha",after=wb.sheets[0])
print(wb.sheets.count)

wb.sheets[0].range('A1').value = "hello world"
wb.sheets[1].range('A1').value = "hello world"
wb2.sheets[0].range('A1').value = "hello world2"

wb.save("test.xlsx")
# wb2.save('test2.xlsx')

wb.close()
app.quit()
print("goodbye world")
