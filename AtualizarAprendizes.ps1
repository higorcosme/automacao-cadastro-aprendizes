$Pasta = "C:\Users\Exemplo\Documentos\APRENDIZES"
$ArquivoBase = "C:\Users\Exemplo\APRENDIZES\APRENDIZES.xlsx"

Write-Host ""
Write-Host "=== INICIANDO PROCESSO ==="
Write-Host ""

# Conversao XLS -> XLSX

$Excel = New-Object -ComObject Excel.Application

$Excel.Visible = $false
$Excel.DisplayAlerts = $false

Get-ChildItem $Pasta |
Where-Object {$_.Extension -eq ".xls"} |
ForEach-Object {

    Write-Host "Convertendo:" $_.Name

    $Workbook = $Excel.Workbooks.Open($_.FullName)

    $NovoArquivo = Join-Path `
        $Pasta `
        ($_.BaseName + ".xlsx")

    $Workbook.SaveAs($NovoArquivo, 51)

    $Workbook.Close($false)

}

$Excel.Quit()

# Atualizacao da base

$Excel = New-Object -ComObject Excel.Application

$Excel.Visible = $false
$Excel.DisplayAlerts = $false

Write-Host ""
Write-Host "Atualizando Power Query..."
Write-Host ""

$Workbook = $Excel.Workbooks.Open($ArquivoBase)

$Workbook.RefreshAll()

Start-Sleep 10

$Workbook.Save()

$Workbook.Close($false)

$Excel.Quit()

Write-Host ""
Write-Host "Pressione ENTER para sair..."
Read-Host