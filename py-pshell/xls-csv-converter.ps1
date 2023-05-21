param(
        [String]$source,
        [String]$dest
    )

$Excel = New-Object -ComObject Excel.Application
$wb = $Excel.Workbooks.Open($source)
$timestamp = Get-Date -Format "dd/MM/yyyy hhmmss"
foreach ($ws in $wb.Worksheets) 
    {
        $ws.SaveAs( $dest + "\" + $timestamp + '.csv', 6)
    }
$Excel.Quit()