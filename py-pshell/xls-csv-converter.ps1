param(
        [String]$source,
        [String]$dest
    )

try{
    $Excel = New-Object -ComObject Excel.Application
}
catch{
    "FATAL | Cannot instantiate Excel" | out-File errors.txt -Append
    throw "Error instantiating Excel object"
}

$timestamp = Get-Date -Format "dd/MM/yyyy hhmmss"

try{
    $wb = $Excel.Workbooks.Open($source)
}
catch{ 
    "Invalid source path | " + $source | out-File errors.txt -Append
}
finally{
    $Excel.Quit()
}

try{
    if (-not (Test-Path $dest)) {
        
        throw [System.IO.FileNotFoundException]
    }
     
    foreach ($ws in $wb.Worksheets) 
        {
            $ws.SaveAs( $dest + "\" + $timestamp + '.csv', 6)
        }
    }   
catch{
        "Invalid destination path | " + $dest | out-File errors.txt -Append
}
finally{
    Write-Host "Closing the Excel instance"
    $Excel.Quit()
}