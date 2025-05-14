<#
.SYNOPSIS
Genera un report delle telecamere Milestone con snapshot, metadati da ShortName e Description, e validazione.

.DESCRIPTION
Estrae i metadati da ShortName e Description, salva snapshot identificati solo da CameraId, genera un file Excel completo con immagini integrate.

.REQUIREMENTS
- MilestonePSTools
- ImportExcel

.NOTES
Creato da Roby – versione finale con struttura semplificata e robusta
#>

# Connessione al Management Server
Connect-ManagementServer -ShowDialog -AcceptEula

# Raccolta dati
$cameraReport = Get-VmsCameraReport -IncludeRetentionInfo
$cameraInfo = Get-VmsCamera

# Crea cartella snapshot
$snapshotFolder = Join-Path $PSScriptRoot "Snapshots_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -Path $snapshotFolder -ItemType Directory -Force | Out-Null

# Funzione parsing Key:Val da stringa
function Get-MetadataValue {
    param (
        [string]$InputString,
        [string]$Key
    )
    $pattern = "(?<![A-Z])$Key\s*:\s*([^;]*)"
    if ($InputString -match $pattern) {
        $value = $matches[1].Trim()
        return $value
    } else {
        return ""
    }
}

# Funzione per inserire immagine in Excel
function Add-ExcelImage {
    param(
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet,
        [System.Drawing.Image]$Image,
        [int]$Row,
        [int]$Column,
        [int]$Width = 150,
        [int]$Height = 100
    )
    $picture = $WorkSheet.Drawings.AddPicture((New-Guid).ToString(), $Image)
    $picture.SetPosition($Row - 1, 5, $Column - 1, 5)
    $picture.SetSize($Width, $Height)
    $WorkSheet.Column($Column).Width = ($Width / 6)
    $WorkSheet.Row($Row).Height = ($Height * 0.75)
}

Write-Host "Elaborazione telecamere in corso..." -ForegroundColor Cyan

$combinedReport = foreach ($cam in $cameraInfo) {

    $reportMatch = $cameraReport | Where-Object { $_.Name -eq $cam.Name } | Select-Object -First 1

    # Snapshot salvato con nome semplificato (solo CameraId)
    try {
        $snapshotPath = Join-Path $snapshotFolder "$($cam.Id).jpg"
        $null = $cam | Get-Snapshot -Behavior GetEnd -Save -Path $snapshotFolder -FileName "$($cam.Id).jpg"
    } catch {
        $snapshotPath = $null
    }

    # Parsing da ShortName
    $procura     = Get-MetadataValue -InputString $cam.ShortName -Key "P"
    $pp          = Get-MetadataValue -InputString $cam.ShortName -Key "PP"
    $registro    = Get-MetadataValue -InputString $cam.ShortName -Key "R"
    $stato       = Get-MetadataValue -InputString $cam.ShortName -Key "S"

    # Parsing da Description
    $reparto       = Get-MetadataValue -InputString $cam.Description -Key "REPARTO"
    $descrizione   = Get-MetadataValue -InputString $cam.Description -Key "DESCRIZIONE"
    $trasmissione  = Get-MetadataValue -InputString $cam.Description -Key "TRASMISSIONE"
    $reset         = Get-MetadataValue -InputString $cam.Description -Key "RESET"
    $esportato     = Get-MetadataValue -InputString $cam.Description -Key "ESPORTATO"
    $note          = Get-MetadataValue -InputString $cam.Description -Key "NOTE"

    # Validità metadati
    $validita = if ($procura -or $pp -or $registro -or $stato) { "✔" } else { "❌" }

    [PSCustomObject]@{
        RecorderName        = $reportMatch.RecorderName
        Procura             = $procura
        Reparto             = $reparto
        'Procedimento Penale' = $pp
        Registro            = $registro
        Name                = $cam.Name
        Descrizione         = $descrizione
        Address             = $reportMatch.Address
        Trasmissione        = $trasmissione
        Reset               = $reset
        Stato               = $stato
        Enabled             = if ($reportMatch) { $true } else { $false }
        Esportato           = $esportato
        Note                = $note
        LastModified        = $cam.LastModified
        MediaDatabaseBegin  = $reportMatch.MediaDatabaseBegin
        MediaDatabaseEnd    = $reportMatch.MediaDatabaseEnd
        UsedSpaceInGB       = $reportMatch.UsedSpaceInGB
        ActualRetentionDays = $reportMatch.ActualRetentionDays
        IsRecording         = $reportMatch.IsRecording
        SnapshotPath        = if ($snapshotPath) { $snapshotPath } else { "No Snapshot" }
        ValiditaMetadati    = $validita
        Id                  = $cam.Id
        ShortName           = $cam.ShortName
        Description         = $cam.Description
    }
}

# Mostra anteprima in griglia
$combinedReport | Out-GridView -Title "Report Completo Telecamere Milestone"

# Salvataggio Excel
$recorderName = ($cameraReport | Select-Object -First 1).RecorderName -replace '[^\w\-]','_'
$fileExcel = Join-Path $PSScriptRoot "${recorderName}_ReportTelecamere_$((Get-Date).ToString('yyyyMMdd_HHmmss')).xlsx"

$excel = $combinedReport | Select RecorderName, Procura, Reparto, 'Procedimento Penale', Registro, Name, Descrizione, Address, Trasmissione, Reset, Stato, Enabled, Esportato, Note, LastModified, MediaDatabaseBegin, MediaDatabaseEnd, UsedSpaceInGB, ActualRetentionDays, IsRecording, ValiditaMetadati |
    Export-Excel -Path $fileExcel -AutoSize -WorksheetName "Telecamere" -PassThru

# Inserimento immagini in colonna 22
$sheet = $excel.Workbook.Worksheets["Telecamere"]
$row = 2
foreach ($cam in $combinedReport) {
    if (Test-Path $cam.SnapshotPath) {
        $image = [System.Drawing.Image]::FromFile($cam.SnapshotPath)
        Add-ExcelImage -WorkSheet $sheet -Image $image -Row $row -Column 22
    }
    $row++
}

Close-ExcelPackage $excel
Write-Host "`n✅ Report generato correttamente: $fileExcel" -ForegroundColor Green
