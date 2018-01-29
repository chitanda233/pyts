import xlwings as xw

app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False

filepath = '怪物模板.xlsx'
workbook1 = app.books.open(filepath)
workbook_out = app.books.add()
sht_out = workbook_out.sheets['sheet1']
nrow_out = sht_out.api.UsedRange.Rows.count


sht = workbook1.sheets["旧版怪"]
# rng_all = sht.range(1, 1).expand().shape
nrow = sht.api.UsedRange.Rows.count
ncol = sht.api.UsedRange.Columns.count

rng_all = sht.range((1, 1), (nrow, ncol))
rng_name = sht.range((1, 1), (nrow, 1))
keyword = "侦察机"
level = "1档"
L2 = []
for name in rng_name:
    if (name.value == keyword):
        L1 = rng_all.rows[name.row - 1]
        L2.append(L1)
L3 = []

print("-----L2-------")
print(L2)
print("------L2---------")

for mon_row in L2:
    if (mon_row(1, 2).value == level):
        start='A'+str(nrow_out)
        print(start)
        sht_out.range(start).value = mon_row.value[4:]
        nrow_out=nrow_out+1

workbook_out.save('输出.xlsx')
workbook_out.close()
workbook1.close()
app.quit()

print("goodbye")
