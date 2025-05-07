# 📄 PowerShell PDF Merger

Este repositório contém dois scripts em PowerShell para mesclar arquivos PDF utilizando módulos distintos:  
- [`mergepdf.ps1`](#1-mergepdfps1): Usa o módulo [PsPdf](https://www.powershellgallery.com/packages/PsPdf)  
- [`mergepdf-with-progress.ps1`](#2-mergepdf-with-progressps1): Usa o módulo [PSWritePDF](https://github.com/EvotecIT/PSWritePDF)

## 📦 Dependências

Antes de utilizar os scripts, instale os módulos necessários com os seguintes comandos:

### Para `mergepdf.ps1`
```powershell
Install-Module -Name PsPdf -Force
```

### Para `mergepdf-with-progress.ps1`
```powershell
Install-Module -Name PSWritePDF -Force
```

---

## 📁 Scripts

### 1. `mergepdf.ps1`

Script simples que:

- Lê todos os arquivos `.pdf` de um diretório
- Ordena por nome
- Mescla todos os arquivos em um único PDF

#### Uso:

```powershell
.\mergepdf.ps1 -InputFolder "C:\Caminho\Para\PDFs" -OutputFile "C:\Destino\Merged.pdf" -SortBy "Name"
```

#### Parâmetros:

- `-InputFolder`: Caminho para o diretório com os PDFs
- `-OutputFile`: Caminho do arquivo PDF final gerado
- `-SortBy`: (opcional) "Name", "CreationTime", "LastWriteTime"
- `-Recurse`: (opcional) Inclui subdiretórios na busca
- `-CoverFile`: (opcional) Arquivo PDF que será utilizado como Capa
- `-$LogFile`: (opcional) LOG com a Lista dos arquivos PDFs unidos

#### Código:

```powershell
#https://www.powershellgallery.com/packages/PsPdf
#Install-Module -Name PsPdf
#mergepdf.ps1
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
```

---

### 2. `mergepdf-with-progress.ps1`

Script mais avançado que:

- Mescla arquivos PDF com **barra de progresso**
- Suporta leitura recursiva de subpastas
- Mostra **tempo estimado restante (ETA)** opcionalmente

#### Uso:

```powershell
.\mergepdf-with-progress.ps1  `
    -InputFolder "C:\Caminho\Para\PDFs" `
    -OutputFile "C:\Destino\Merged.pdf" `
    -Recurse `
    -SortBy "Name" `
    -CoverFile "C:\Caminho\Para\Cover.pdf" `
    -LogFile "C:\Caminho\Para\merge-log.txt" `
    -ShowETA
```

#### Parâmetros:

- `-InputFolder`: Caminho para o diretório com os PDFs
- `-OutputFile`: Caminho do arquivo PDF final gerado
- `-SortBy`: (opcional) "Name", "CreationTime", "LastWriteTime"
- `-Recurse`: (opcional) Inclui subdiretórios na busca
- `-CoverFile`: (opcional) Arquivo PDF que será utilizado como Capa
- `-$LogFile`: (opcional) LOG com a Lista dos arquivos PDFs unidos
- `-ShowETA`: (opcional) Exibe tempo restante na barra de progresso

#### Código:

```powershell
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
```
---

## 📌 Observações

- Scripts compatíveis com PowerShell 5.1 ou superior.
- Recomendado executar como Administrador caso o destino exija permissões elevadas.

---

## 🧑‍💻 Autor

**Marinaldo de Jesus**  
Scripts desenvolvidos para automatizar a preparação de documentos acadêmicos.

---

## 📄 Licença

Distribuído sob a licença MIT. Veja `LICENSE.md` para mais detalhes.
