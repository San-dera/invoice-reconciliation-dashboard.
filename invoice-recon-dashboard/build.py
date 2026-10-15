import json, openpyxl
wb = openpyxl.load_workbook('data/111_-_Combining_Text_Data_analysed.xlsx', data_only=True)
ws = wb['MC Invoice Report']
rows = []
for r in ws.iter_rows(min_row=5, max_col=13, values_only=True):
    if not r[0]: continue
    rows.append(dict(ref=r[0], supplier=r[1], inv=r[3].date().isoformat(), due=r[4].date().isoformat(),
                     paid=r[5].date().isoformat(), po=r[7], loc=r[8], amt=r[9], month=r[10], days=r[11]))
ra = wb['Recon Analysis']
recon = dict(stmt=ra['B3'].value, sys=ra['B4'].value, diff=ra['B5'].value, credits_n=ra['B13'].value,
             credits=ra['B14'].value, naive=ra['B15'].value, over=ra['B16'].value, dup=ra['B17'].value)
data = dict(rows=rows, recon=recon, rate=ws['M2'].value)
open('index.html', 'w').write(open('template.html').read().replace('__DATA__', json.dumps(data)))
print(len(rows), recon)
