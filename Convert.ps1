Get-ChildItem .\* -Include *.xls* | ForEach-Object {

    $excelFile = $_.FullName
    $excelFileName = $_.Name.Split(".")[0]
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false

    $wb = $E.Workbooks.Open($excelFile)

    foreach ($ws in $wb.Worksheets)
    {
        $n = $excelFileName + "_" + $ws.Name
        $path = $(Get-Location).Path + "\converted\"
        If(!(test-path $path))
        {
            New-Item -ItemType Directory -Force -Path $path
        }
        $newFile = $path + $n + ".csv"
        Write-host "Convertendo a planilha: " + $newFile
        $ws.SaveAs($newFile, 62)
        Write-Host $n
    }

    $E.Quit()

}