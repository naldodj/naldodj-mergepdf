#https://github.com/EvotecIT/PSWritePDF
#Install-Module PSWritePDF -Force
#file:mergepdf-with-progress.ps1
#autor:Marinaldo de Jesus
param (
    [Parameter(Mandatory = $true)]
    [string[]]$InputFile,

    [Parameter(Mandatory = $true)]
    [string]$OutputFile,

    [ValidateSet("Name", "CreationTime", "LastWriteTime")]
    [string]$SortBy = "Name",

    [switch]$Recurse,

    [string]$CoverFile,

    [string]$LogFile,

    [switch]$ShowETA

)

Import-Module PSWritePDF -Force

# Coletar arquivos
$Files = @()
foreach ($entry in $InputFile) {
    if (Test-Path $entry -PathType Container) {
        $Files += Get-ChildItem $entry -Filter "*.pdf" -File -Recurse:($Recurse.IsPresent)
    } elseif (Test-Path $entry -PathType Leaf -and $entry.ToLower().EndsWith(".pdf")) {
        $Files += Get-Item $entry
    }
}
if ($Files.Count -eq 0) {
    Write-Warning "Nenhum arquivo PDF válido encontrado."
    return
}

# Ordenar conforme critério
$Files = $Files | Sort-Object $SortBy

# Inserir capa se especificada
if ($CoverFile -and (Test-Path $CoverFile -PathType Leaf)) {
    $Files = ,(Get-Item $CoverFile) + $Files
}

# Exportar log se solicitado
if ($LogFile) {
    $Files | Select-Object -ExpandProperty FullName | Out-File -FilePath $LogFile -Encoding UTF8
    Write-Host "Log de arquivos gravado em: $LogFile" -ForegroundColor Cyan
}

# Preparar mesclagem incremental
$tempPath  = [System.IO.Path]::GetDirectoryName($OutputFile)
$FinalTemp = Join-Path $tempPath "MergedFile0.pdf"
Copy-Item $Files[0].FullName $FinalTemp -Force

# Tempo para ETA
if ($ShowETA) {
    $startTime = Get-Date
}

for ($i = 1; $i -lt $Files.Count; $i++) {
    $source = $Files[$i].FullName
    $newTemp = Join-Path $tempPath "MergedFile$i.pdf"

    # Cálculo de ETA
    if ($ShowETA) {
        $elapsed    = (Get-Date) - $startTime
        $avgPerFile = $elapsed.TotalSeconds / $i
        $remaining  = [int]( $avgPerFile * ($Files.Count - $i) )

        Write-Progress `
            -Activity "Mesclando PDFs" `
            -Status   "Adicionando: $([IO.Path]::GetFileName($source))" `
            -PercentComplete (($i / $Files.Count) * 100) `
            -SecondsRemaining ($ShowETA ? $remaining : $null)

    } else {

        Write-Progress `
            -Activity "Mesclando PDFs" `
            -Status   "Adicionando: $([IO.Path]::GetFileName($source))" `
            -PercentComplete (($i / $Files.Count) * 100)

    }

    Merge-PDF -InputFile $FinalTemp, $source -OutputFile $newTemp
    Remove-Item -LiteralPath $FinalTemp -Force
    $FinalTemp = $newTemp
}

# Mover para arquivo final
if (Test-Path $OutputFile) {
    Remove-Item $OutputFile -Force
}
Move-Item -LiteralPath $FinalTemp -Destination $OutputFile

Write-Progress -Activity "Mesclando PDFs" -Completed -Status "Concluído!"
Write-Host "PDF final salvo em: $OutputFile" -ForegroundColor Green
