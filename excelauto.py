import win32com.client as win32
excel=win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible=True
wb=excel.Workbooks.Add()
ws = wb.Worksheets('Sheet1')
ws.Name='GeneratedData'
ws.Cells(1,1).Value= 'Data Analysis'
ws.Cells(1,1).Font.Size=15
ws.Cells(1,1).Font.Bold=True
for i in range (1,6) :
    ws.Cells(2,i).Value=i
ws.Range(ws.Cells(3,1),ws.Cells(3,5)).Value=[10,20,30,40,50]
ws.Range("A4:E4").Value= [i+100 for i in range(1,6)]

for i in range (1,6) :
    ws.Cells(5,i).Value= ws.Cells(2,i).Value + ws.Cells(3,i).Value + ws.Cells(4,i).Value
    ws.Cells(5,i).Font.Size=15
    ws.Cells(5,i).Font.Bold=True