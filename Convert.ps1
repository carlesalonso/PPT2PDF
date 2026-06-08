$scriptRoot   = Split-Path -Parent $MyInvocation.MyCommand.Path
$inputFolder  = Join-Path $scriptRoot "input"
$outputFolder = Join-Path $scriptRoot "output"

# Constantes COM para no depender de ensamblados Office en PowerShell.
# msoFalse (MsoTriState): valor 0 = False. Se usa en Open() para abrir sin
# ventana, sin solo-lectura y sin añadir la presentacion a la lista reciente.
$msoFalse    = 0
# ppSaveAsPDF (PpSaveAsFileType): valor 32 = exportar/guardar como PDF.
$ppSaveAsPDF = 32

if (-not (Test-Path $outputFolder)) { New-Item -ItemType Directory -Path $outputFolder | Out-Null }

$pptFiles = Get-ChildItem -Path $inputFolder -Filter *.ppt* -File -Recurse

try {
    $pptApp = New-Object -ComObject PowerPoint.Application -ErrorAction Stop
}
catch {
    throw "No se pudo iniciar PowerPoint vía COM. Verifica instalación de Microsoft PowerPoint y permisos de la sesión. Detalle: $($_.Exception.Message)"
}

foreach ($f in $pptFiles) {
    try {
        $full = $f.FullName
        $base = [System.IO.Path]::GetFileNameWithoutExtension($full)
        $pdf  = Join-Path $outputFolder ("$base.pdf")
        $pres = $pptApp.Presentations.Open(
            $full,
            $msoFalse,
            $msoFalse,
            $msoFalse
        )
        $pres.SaveAs($pdf, $ppSaveAsPDF)
        $pres.Close()
        Write-Host "Convertido: $($f.Name) -> $pdf"
    } catch {
        Write-Warning "Error con $($f.Name): $_"
    }
}

# Liberar recursos COM de PowerPoint.

if ($null -ne $pptApp) {
    $pptApp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptApp) | Out-Null
}
