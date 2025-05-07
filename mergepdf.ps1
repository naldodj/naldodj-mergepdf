#https://www.powershellgallery.com/packages/PsPdf
#Install-Module -Name PsPdf
#file:mergepdf.ps1
#autor:Marinaldo de Jesus
param(
    [Parameter(Mandatory)]
    [string]$InputFolder,

    [Parameter(Mandatory)]
    [string]$OutputFile,

    [ValidateSet("Name", "CreationTime", "LastWriteTime")]
    [string]$SortBy = "Name",

    [switch]$Recurse,

    [string]$CoverFile,

    [string]$LogFile

)

Import-Module PsPdf -Force

# Coletar arquivos
$Files = @()
foreach ($entry in $InputFolder) {
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

Merge-Pdf -InputFile $Files -OutputFile $OutputFile
Write-Host "PDF final salvo em: $OutputFile" -ForegroundColor Green
