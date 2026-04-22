# =============================================================================
# Wells.pt Price Scraper
# Recolhe precos e descontos de Chupetas, Biberoes e Bombas Tira Leite
# Nao requer Python, Node.js ou Claude - apenas PowerShell nativo
# Agendar via: powershell -ExecutionPolicy Bypass -File agendar_tarefa.ps1
# =============================================================================

param(
    [string]$OutputDir = "$PSScriptRoot\data",
    [switch]$VerboseMode
)

$ErrorActionPreference = "Continue"

# ---------------------------------------------------------------------------
# Configuracao das categorias
# ---------------------------------------------------------------------------
$CATEGORIES = @(
    @{ Name = "Chupetas";           CgId = "bebe-chuchas-chuchas";     Total = 250 },
    @{ Name = "Biberoes";           CgId = "bebe-alimentacao-biberoes"; Total = 150 },
    @{ Name = "Bombas_Tira_Leite";  CgId = "mama-bombas-leite";        Total = 60  }
)

$BASE_HOST = "https://wells.pt"
$BASE_PATH = "/on/demandware.store/Sites-Wells-Site/pt_PT/Search-ShowAjax"
$PAGE_SIZE = 100

$HEADERS = @{
    "User-Agent"      = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    "Accept"          = "text/html,application/xhtml+xml,application/xml;q=0.9"
    "Referer"         = "https://wells.pt/"
    "Accept-Language" = "pt-PT,pt;q=0.9"
}

# ---------------------------------------------------------------------------
# Estrutura de pastas
# ---------------------------------------------------------------------------
$DirRecente   = Join-Path $OutputDir "recente"      # ultimos 30 dias
$DirArquivo   = Join-Path $OutputDir "arquivo"      # meses anteriores (zip)
$DirHistorico = Join-Path $OutputDir "historico"    # CSV acumulado
$DirLogs      = Join-Path $OutputDir "logs"         # logs com rotacao mensal

foreach ($d in @($DirRecente, $DirArquivo, $DirHistorico, $DirLogs)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

# Caminhos dos ficheiros de hoje
$Date      = Get-Date -Format "yyyy-MM-dd"
$YearMonth = Get-Date -Format "yyyy-MM"
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$CsvDaily  = Join-Path $DirRecente  ("wells_" + $Date + ".csv")
$ReportFile= Join-Path $DirRecente  ("relatorio_" + $Date + ".html")
$CsvMaster = Join-Path $DirHistorico "wells_historico.csv"
$LogFile   = Join-Path $DirLogs     ("scraper_" + $YearMonth + ".log")

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[" + $Timestamp + "] [" + $Level + "] " + $Message
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    if ($VerboseMode) { Write-Host $line }
}

# ---------------------------------------------------------------------------
# Gestao de ficheiros: arquivar o que tem mais de 30 dias
# ---------------------------------------------------------------------------
function Manage-Files {
    # Arquivar ficheiros cujo mes (no nome) seja anterior ao mes atual
    $currentYM = Get-Date -Format "yyyy-MM"

    # Migrar ficheiros diarios antigos da raiz data\ (execucoes anteriores sem estrutura)
    # Exclui o historico e o log que tem destinos proprios
    Get-ChildItem -Path $OutputDir -File | Where-Object {
        $_.Name -match '^(wells_\d{4}-\d{2}-\d{2}|relatorio_\d{4}-\d{2}-\d{2})'
    } | ForEach-Object {
        $dest = Join-Path $DirRecente $_.Name
        if (-not (Test-Path $dest)) { Move-Item $_.FullName $dest }
    }

    # Mover ficheiros de meses anteriores de recente\ para arquivo\YYYY-MM\
    Get-ChildItem -Path $DirRecente -File | Where-Object {
        # Extrair YYYY-MM do nome; se o mes for anterior ao atual, arquivar
        if ($_.Name -match '(\d{4}-\d{2})-\d{2}') { $Matches[1] -lt $currentYM } else { $false }
    } | ForEach-Object {
        # Extrair YYYY-MM do nome do ficheiro (ex: wells_2026-04-06.csv -> 2026-04)
        $ym = $null
        if ($_.Name -match '(\d{4}-\d{2})-\d{2}') { $ym = $Matches[1] }
        if (-not $ym) { $ym = $_.LastWriteTime.ToString("yyyy-MM") }

        $destDir = Join-Path $DirArquivo $ym
        if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }

        Move-Item $_.FullName (Join-Path $destDir $_.Name) -Force
        Write-Log ("Arquivado: " + $_.Name + " -> arquivo\" + $ym + "\")
    }

    # Compactar em ZIP pastas de arquivo de meses ja encerrados (todos exceto o atual)
    Get-ChildItem -Path $DirArquivo -Directory | Where-Object {
        $_.Name -match '^\d{4}-\d{2}$' -and $_.Name -lt $currentYM
    } | ForEach-Object {
        $zipPath = Join-Path $DirArquivo ($_.Name + ".zip")
        if (-not (Test-Path $zipPath)) {
            try {
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                [System.IO.Compression.ZipFile]::CreateFromDirectory($_.FullName, $zipPath)
                Remove-Item $_.FullName -Recurse -Force
                Write-Log ("Compactado: arquivo\" + $_.Name + "\ -> " + $_.Name + ".zip")
            } catch {
                Write-Log ("Erro ao compactar " + $_.Name + ": " + $_) "ERROR"
            }
        }
    }

    # Migrar wells_historico.csv da raiz para historico\ (caso exista da versao antiga)
    $oldMaster = Join-Path $OutputDir "wells_historico.csv"
    if ((Test-Path $oldMaster) -and -not (Test-Path $CsvMaster)) {
        Move-Item $oldMaster $CsvMaster
        Write-Log "Migrado wells_historico.csv para historico\"
    }

    # Migrar scraper.log da raiz para logs\ (caso exista da versao antiga)
    $oldLog = Join-Path $OutputDir "scraper.log"
    if (Test-Path $oldLog) {
        $archiveLog = Join-Path $DirLogs "scraper_migrado.log"
        Get-Content $oldLog | Add-Content $archiveLog -Encoding UTF8
        Remove-Item $oldLog -Force
        Write-Log "Migrado scraper.log para logs\scraper_migrado.log"
    }

    Write-Log "Gestao de ficheiros concluida"
}

# ---------------------------------------------------------------------------
# Decode HTML entities
# ---------------------------------------------------------------------------
function Decode-HtmlEntities {
    param([string]$Text)
    $Text = $Text -replace '&quot;',   '"'
    $Text = $Text -replace '&amp;',    '&'
    $Text = $Text -replace '&lt;',     '<'
    $Text = $Text -replace '&gt;',     '>'
    $Text = $Text -replace '&eacute;', 'e'
    $Text = $Text -replace '&atilde;', 'a'
    $Text = $Text -replace '&oacute;', 'o'
    $Text = $Text -replace '&aacute;', 'a'
    $Text = $Text -replace '&uacute;', 'u'
    $Text = $Text -replace '&iacute;', 'i'
    $Text = $Text -replace '&ccedil;', 'c'
    return $Text
}

# ---------------------------------------------------------------------------
# Parser: usa data-product-tile-impression (JSON embutido no HTML)
# ---------------------------------------------------------------------------
function Parse-Products {
    param([string]$Html, [string]$Category)

    $products = New-Object System.Collections.ArrayList
    $tilePattern = 'data-product-tile-impression="(\{[^"]{20,3000}\})"'
    $tiles = [regex]::Matches($Html, $tilePattern)

    foreach ($tile in $tiles) {
        $jsonRaw = Decode-HtmlEntities $tile.Groups[1].Value

        $nameMatch         = [regex]::Match($jsonRaw, '"name"\s*:\s*"([^"]+)"')
        $brandMatch        = [regex]::Match($jsonRaw, '"brand"\s*:\s*"([^"]*)"')
        $idMatch           = [regex]::Match($jsonRaw, '"id"\s*:\s*"([^"]+)"')
        $priceMatch        = [regex]::Match($jsonRaw, '"price"\s*:\s*([\d.]+)')
        $pvpMatch          = [regex]::Match($jsonRaw, '"pvp"\s*:\s*([\d.]+)')
        $urlMatch          = [regex]::Match($jsonRaw, '"url"\s*:\s*"([^"]+)"')
        $availableMatch    = [regex]::Match($jsonRaw, '"available"\s*:\s*(true|false)')
        $readyToOrderMatch = [regex]::Match($jsonRaw, '"readyToOrder"\s*:\s*(true|false)')
        $notifyMatch       = [regex]::Match($jsonRaw, '"notify"\s*:\s*(true|false)')

        if (-not $nameMatch.Success -or -not $priceMatch.Success) { continue }

        $name   = $nameMatch.Groups[1].Value.Trim()
        $brand  = if ($brandMatch.Success) { $brandMatch.Groups[1].Value.Trim() } else { "" }
        $prodId = if ($idMatch.Success)    { $idMatch.Groups[1].Value }           else { "" }

        # No JSON da Wells: "price" = PVPR (original), "pvp" = preco atual
        $pvprRaw  = [decimal]$priceMatch.Groups[1].Value
        $priceRaw = if ($pvpMatch.Success) { [decimal]$pvpMatch.Groups[1].Value } else { $pvprRaw }
        $price    = $priceRaw
        $pvpr     = if ($pvprRaw -ne $priceRaw) { $pvprRaw } else { $null }

        # Stock: campos do JSON do tile (por ordem de fiabilidade)
        # available=false -> sem stock | readyToOrder=false -> sem stock | notify=true -> sem stock
        $stock = "Disponivel"
        if ($availableMatch.Success    -and $availableMatch.Groups[1].Value    -eq "false") { $stock = "Sem Stock" }
        elseif ($readyToOrderMatch.Success -and $readyToOrderMatch.Groups[1].Value -eq "false") { $stock = "Sem Stock" }
        elseif ($notifyMatch.Success   -and $notifyMatch.Groups[1].Value       -eq "true")  { $stock = "Sem Stock" }

        $discount = $null
        if ($null -ne $pvpr -and $pvpr -gt $price -and $pvpr -gt 0) {
            $discount = [math]::Round((($pvpr - $price) / $pvpr) * 100)
        }

        $saving = $null
        if ($null -ne $pvpr -and $pvpr -gt $price) {
            $saving = [math]::Round($pvpr - $price, 2)
        }

        $pos    = $tile.Index
        $window = $Html.Substring([Math]::Max(0, $pos - 200), [Math]::Min(2000, $Html.Length - [Math]::Max(0, $pos - 200)))
        $labels = New-Object System.Collections.ArrayList
        if ($window -match 'Best\s*Seller')      { [void]$labels.Add("Best Seller") }
        if ($window -match 'Exclusivo\s*Online') { [void]$labels.Add("Exclusivo Online") }
        if ($window -match '"novo"' -or $window -match '>Novo<') { [void]$labels.Add("Novo") }

        $tileUrlMatch = [regex]::Match($window, 'data-product-tile-url="([^"]+)"')
        $prodUrl = if ($tileUrlMatch.Success) {
            $BASE_HOST + $tileUrlMatch.Groups[1].Value
        } elseif ($urlMatch.Success -and $urlMatch.Groups[1].Value -notmatch "\.html$|bebe-mama") {
            $urlMatch.Groups[1].Value
        } else {
            $BASE_HOST + "/" + $prodId + ".html"
        }

        [void]$products.Add([PSCustomObject]@{
            Data          = $Date
            Hora          = (Get-Date -Format "HH:mm:ss")
            Categoria     = $Category
            ProdID        = $prodId
            Marca         = $brand
            Produto       = $name
            Preco         = $price
            PVPR          = $pvpr
            Desconto_Pct  = $discount
            Poupanca_Euro = $saving
            Destaque      = ($labels -join "; ")
            Stock         = $stock
            URL           = $prodUrl
        })
    }

    return $products
}

# ---------------------------------------------------------------------------
# Buscar uma pagina da API
# ---------------------------------------------------------------------------
function Fetch-Page {
    param([string]$CgId, [int]$Start)

    $qs  = "cgid=" + $CgId + "&srule=best-matches&start=" + $Start + "&sz=" + $PAGE_SIZE
    $url = $BASE_HOST + $BASE_PATH + "?" + $qs

    try {
        $resp = Invoke-WebRequest -Uri $url -Headers $HEADERS -UseBasicParsing -TimeoutSec 30
        return $resp.Content
    } catch {
        Write-Log ("Erro HTTP: " + $url + " | " + $_) "ERROR"
        return $null
    }
}

# ===========================================================================
# EXECUCAO PRINCIPAL
# ===========================================================================
Write-Log "=== Inicio do scraping wells.pt ==="

# 1. Gerir ficheiros antigos antes de criar novos
Manage-Files

# 2. Scraping
$allProducts = New-Object System.Collections.ArrayList
$summary     = @{}

foreach ($cat in $CATEGORIES) {
    Write-Log ("A processar: " + $cat.Name)
    $catProducts = New-Object System.Collections.ArrayList
    $start = 0
    $pages = 0

    do {
        $html = Fetch-Page -CgId $cat.CgId -Start $start
        if ($null -eq $html -or $html.Length -lt 100) { break }

        # DEBUG: log JSON do 1o tile da 1a pagina para diagnostico de stock
        if ($start -eq 0 -and $pages -eq 0) {
            $dbgM = [regex]::Match($html, 'data-product-tile-impression="(\{[^"]{20,3000}\})"')
            if ($dbgM.Success) {
                $dbgJ = $dbgM.Groups[1].Value -replace '&quot;','"' -replace '&amp;','&'
                Write-Log ("[DEBUG] Tile JSON (" + $cat.Name + "): " + $dbgJ.Substring(0, [Math]::Min(500, $dbgJ.Length)))
            }
        }

        $parsed = Parse-Products -Html $html -Category $cat.Name
        if ($parsed.Count -eq 0) { break }

        foreach ($p in $parsed) { [void]$catProducts.Add($p) }
        Write-Log ("  Pag " + ($pages + 1) + ": +" + $parsed.Count + " produtos (total: " + $catProducts.Count + ")")

        $start += $PAGE_SIZE
        $pages++
        Start-Sleep -Milliseconds 600

    } while ($parsed.Count -eq $PAGE_SIZE -and $start -lt $cat.Total)

    foreach ($p in $catProducts) { [void]$allProducts.Add($p) }

    $withDiscount = @($catProducts | Where-Object { $null -ne $_.Desconto_Pct }).Count
    $outOfStock   = @($catProducts | Where-Object { $_.Stock -eq "Sem Stock" }).Count
    $maxDisc      = if ($withDiscount -gt 0) {
        ($catProducts | Where-Object { $null -ne $_.Desconto_Pct } | Measure-Object -Property Desconto_Pct -Maximum).Maximum
    } else { 0 }

    $summary[$cat.Name] = @{
        Total       = $catProducts.Count
        ComDesconto = $withDiscount
        SemStock    = $outOfStock
        MaxDesconto = $maxDisc
        PrecoMin    = ($catProducts | Measure-Object -Property Preco -Minimum).Minimum
        PrecoMax    = ($catProducts | Measure-Object -Property Preco -Maximum).Maximum
    }

    Write-Log ("  OK: " + $catProducts.Count + " produtos | " + $withDiscount + " c/ desconto | max " + $maxDisc + "%")
}

# 3. Exportar CSV diario em recente\
if ($allProducts.Count -gt 0) {
    $allProducts | Export-Csv -Path $CsvDaily -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Log ("CSV diario: " + $CsvDaily + " (" + $allProducts.Count + " linhas)")

    # Adicionar ao historico acumulado em historico\
    if (-not (Test-Path $CsvMaster)) {
        $allProducts | Export-Csv -Path $CsvMaster -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    } else {
        $allProducts | ConvertTo-Csv -NoTypeInformation -Delimiter ";" |
            Select-Object -Skip 1 |
            Add-Content -Path $CsvMaster -Encoding UTF8
    }
    Write-Log ("Historico: " + $CsvMaster)
} else {
    Write-Log "AVISO: Nenhum produto extraido!" "WARN"
}

# 4. Relatorio HTML em recente\
$rowsHtml = ($allProducts | ForEach-Object {
    $dCell   = if ($null -ne $_.Desconto_Pct) { "<td style='color:#c0392b;font-weight:bold'>-" + $_.Desconto_Pct + "%</td>" } else { "<td>-</td>" }
    $pvCell  = if ($null -ne $_.PVPR)         { "<td>" + $_.PVPR + "</td>" }                                                   else { "<td>-</td>" }
    $sCell   = if ($null -ne $_.Poupanca_Euro){ "<td style='color:#27ae60'>" + $_.Poupanca_Euro + "</td>" }                    else { "<td>-</td>" }
    $stCell  = if ($_.Stock -eq "Sem Stock")  { "<td style='color:#e74c3c'>Sem Stock</td>" }                                   else { "<td style='color:#27ae60'>Disponivel</td>" }
    $dstCell = if ($_.Destaque) { "<td><span style='background:#f39c12;color:#fff;padding:2px 6px;border-radius:3px;font-size:11px'>" + $_.Destaque + "</span></td>" } else { "<td></td>" }
    "<tr><td>" + $_.Categoria + "</td><td>" + $_.Marca + "</td><td><a href='" + $_.URL + "' target='_blank'>" + $_.Produto + "</a></td><td>" + $_.Preco + "</td>" + $pvCell + $dCell + $sCell + $stCell + $dstCell + "</tr>"
}) -join "`n"

$summHtml = ($CATEGORIES | ForEach-Object {
    $s = $summary[$_.Name]
    if ($s) { "<tr><td><strong>" + $_.Name + "</strong></td><td>" + $s.Total + "</td><td>" + $s.ComDesconto + "</td><td>" + $s.SemStock + "</td><td>" + $s.MaxDesconto + "%</td><td>" + $s.PrecoMin + "</td><td>" + $s.PrecoMax + "</td></tr>" }
}) -join "`n"

# Links de navegacao rapida para outros relatorios
$otherReports = Get-ChildItem -Path $DirRecente -Filter "relatorio_*.html" |
    Where-Object { $_.Name -ne ("relatorio_" + $Date + ".html") } |
    Sort-Object Name -Descending | Select-Object -First 10 |
    ForEach-Object { "<a href='" + $_.Name + "' style='margin-right:12px'>" + ([regex]::Match($_.Name,'(\d{4}-\d{2}-\d{2})').Groups[1].Value) + "</a>" }
$navHtml = if ($otherReports) { "<p class='meta'>Relatorios anteriores: " + ($otherReports -join "") + "</p>" } else { "" }

$reportHtml  = "<!DOCTYPE html><html lang='pt'><head><meta charset='UTF-8'>"
$reportHtml += "<title>Wells Precos " + $Date + "</title>"
$reportHtml += "<style>body{font-family:Arial,sans-serif;margin:20px;background:#f5f5f5}"
$reportHtml += "h1{color:#1a5276}h2{color:#2874a6;margin-top:30px}"
$reportHtml += "table{border-collapse:collapse;width:100%;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.1);margin-bottom:30px}"
$reportHtml += "th{background:#2874a6;color:#fff;padding:10px 8px;text-align:left;font-size:13px}"
$reportHtml += "td{padding:8px;border-bottom:1px solid #eee;font-size:12px}"
$reportHtml += "tr:hover{background:#eaf2ff}a{color:#2980b9}.meta{color:#777;font-size:12px}"
$reportHtml += "</style></head><body>"
$reportHtml += "<h1>Wells.pt - Precos e Descontos</h1>"
$reportHtml += "<p class='meta'>Gerado em " + $Timestamp + " | Total: " + $allProducts.Count + " produtos</p>"
$reportHtml += $navHtml
$reportHtml += "<h2>Resumo por Categoria</h2>"
$reportHtml += "<table><tr><th>Categoria</th><th>Total</th><th>Com Desconto</th><th>Sem Stock</th><th>Max Desconto</th><th>Preco Min</th><th>Preco Max</th></tr>"
$reportHtml += $summHtml + "</table>"
$reportHtml += "<h2>Todos os Produtos</h2>"
$reportHtml += "<table><tr><th>Categoria</th><th>Marca</th><th>Produto</th><th>Preco</th><th>PVPR</th><th>Desconto</th><th>Poupanca</th><th>Stock</th><th>Destaque</th></tr>"
$reportHtml += $rowsHtml + "</table></body></html>"

$reportHtml | Out-File -FilePath $ReportFile -Encoding UTF8
Write-Log ("Relatorio HTML: " + $ReportFile)

# 5. Output final
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ("  Wells Scraper - " + $Date) -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
foreach ($cat in $CATEGORIES) {
    $s = $summary[$cat.Name]
    if ($s) {
        Write-Host ("  " + $cat.Name + ": " + $s.Total + " produtos | " + $s.ComDesconto + " c/ desconto | max " + $s.MaxDesconto + "%") -ForegroundColor Green
    }
}
Write-Host ""
Write-Host ("  Hoje:      recente\wells_" + $Date + ".csv") -ForegroundColor Yellow
Write-Host ("  Relatorio: recente\relatorio_" + $Date + ".html") -ForegroundColor Yellow
Write-Host ("  Historico: historico\wells_historico.csv") -ForegroundColor Yellow
Write-Host ("  Log:       logs\scraper_" + $YearMonth + ".log") -ForegroundColor Yellow
Write-Host ""
Write-Log ("=== Concluido. Total: " + $allProducts.Count + " produtos ===")

# 6. Detetar produtos novos e enviar alerta por email
$resendKey = $env:RESEND_API_KEY
if ($resendKey) {
    try {
        # IDs ja conhecidos no historico (dias anteriores)
        $HistoricoFile = Join-Path $OutputDir "historico\wells_historico.csv"
        $idsConhecidos = @{}
        if (Test-Path $HistoricoFile) {
            $historico = Import-Csv -Path $HistoricoFile -Delimiter ";" -Encoding UTF8
            foreach ($row in $historico) {
                if ($row.Data -ne $Date) { $idsConhecidos[$row.ProdID] = $true }
            }
        }

        # Modo de teste: forcar 3 produtos como "novos"
        $modoTeste = $env:TESTE_EMAIL -eq "true"
        if ($modoTeste) {
            $novos = $allProducts | Select-Object -First 3
            Write-Log "MODO TESTE: a enviar email com 3 produtos de exemplo"
        } else {
            $novos = $allProducts | Where-Object { -not $idsConhecidos.ContainsKey($_.ProdID) }
        }

        if ($novos.Count -gt 0) {
            Write-Log ("Produtos novos encontrados: " + $novos.Count)

            # Construir email HTML
            $linhas = $novos | ForEach-Object {
                $preco = $_.Preco.ToString("0.00", [System.Globalization.CultureInfo]::InvariantCulture)
                $pvpr  = if ($_.PVPR) { " <span style='color:#999;text-decoration:line-through'>&euro;" + $_.PVPR.ToString("0.00", [System.Globalization.CultureInfo]::InvariantCulture) + "</span>" } else { "" }
                $desc  = if ($_.Desconto_Pct) { " <span style='color:#e74c3c;font-weight:bold'>-" + $_.Desconto_Pct + "%</span>" } else { "" }
                "<tr><td style='padding:8px;border-bottom:1px solid #eee'>" + $_.Categoria + "</td>" +
                "<td style='padding:8px;border-bottom:1px solid #eee'>" + $_.Marca + "</td>" +
                "<td style='padding:8px;border-bottom:1px solid #eee'><a href='" + $_.URL + "'>" + $_.Produto + "</a></td>" +
                "<td style='padding:8px;border-bottom:1px solid #eee'>&euro;" + $preco + $pvpr + $desc + "</td></tr>"
            }

            $emailHtml = @"
<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto">
  <h2 style="color:#1a5276">Wells.pt — $($novos.Count) produto(s) novo(s) detetado(s)</h2>
  <p style="color:#555">Data: $Date</p>
  <table style="width:100%;border-collapse:collapse">
    <thead>
      <tr style="background:#1a5276;color:#fff">
        <th style="padding:8px;text-align:left">Categoria</th>
        <th style="padding:8px;text-align:left">Marca</th>
        <th style="padding:8px;text-align:left">Produto</th>
        <th style="padding:8px;text-align:left">Preco</th>
      </tr>
    </thead>
    <tbody>
      $($linhas -join "`n")
    </tbody>
  </table>
  <p style="margin-top:16px"><a href="https://tv3nda.github.io/Artsana">Ver dashboard completo</a></p>
</div>
"@

            $body = @{
                from    = "Wells Scraper <onboarding@resend.dev>"
                to      = @("tomasvenda@hotmail.com")
                subject = "Wells.pt — $($novos.Count) produto(s) novo(s) em $Date"
                html    = $emailHtml
            } | ConvertTo-Json -Depth 3

            $headers = @{
                "Authorization" = "Bearer $resendKey"
                "Content-Type"  = "application/json"
            }

            $resp = Invoke-WebRequest -Uri "https://api.resend.com/emails" -Method POST -Headers $headers -Body $body -UseBasicParsing
            Write-Log ("Email enviado. Status: " + $resp.StatusCode)
            Write-Host ("  Email enviado: " + $novos.Count + " produtos novos") -ForegroundColor Green
        } else {
            Write-Log "Sem produtos novos hoje."
            Write-Host "  Sem produtos novos hoje." -ForegroundColor Gray
        }
    } catch {
        Write-Log ("AVISO: Alerta email falhou - " + $_.Exception.Message)
        Write-Host ("  AVISO: Email nao enviado - " + $_.Exception.Message) -ForegroundColor Red
    }
}

# 8. Gerar dashboard interativo
$DashboardScript = Join-Path $PSScriptRoot "gerar_dashboard.ps1"
if (Test-Path $DashboardScript) {
    Write-Host "  A gerar dashboard interativo..." -ForegroundColor Cyan
    Write-Log "A chamar gerar_dashboard.ps1..."
    try {
        & $DashboardScript -DataDir $OutputDir
        Write-Host ("  Dashboard: " + $OutputDir + "\dashboard.html") -ForegroundColor Yellow
        Write-Log "Dashboard gerado com sucesso."
    } catch {
        Write-Log ("AVISO: Dashboard falhou - " + $_.Exception.Message)
        Write-Host ("  AVISO: Dashboard nao gerado - " + $_.Exception.Message) -ForegroundColor Red
    }
    Write-Host ""
}
