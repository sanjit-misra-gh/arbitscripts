$Excel = New-Object -ComObject Excel.Application
$wb = $Excel.Workbooks.Open('C:\Users\sanji\Downloads\testxls.xls')
$timestamp = Get-Date -Format "dd/MM/yyyy hhmmss"
foreach ($ws in $wb.Worksheets) 
    {
        $ws.SaveAs('C:\Users\sanji\Downloads\' + $timestamp + '.csv', 6)
    }
$Excel.Quit()