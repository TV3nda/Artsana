# =============================================================================
# Wells Dashboard Generator
# Gera um dashboard HTML interativo a partir do historico CSV
# Chamado automaticamente pelo wells_scraper.ps1
# Pode tambem ser corrido manualmente:
#   powershell -ExecutionPolicy Bypass -File gerar_dashboard.ps1
# =============================================================================

param(
    [string]$DataDir      = "$PSScriptRoot\data",
    [string]$DashboardOut = "",
    [int]$DiasHistorico   = 90    # quantos dias de historico embutir no dashboard
)

$DbPath    = Join-Path $DataDir "historico\wells.db"
if (-not $DashboardOut) { $DashboardOut = Join-Path $DataDir "dashboard.html" }
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"

if (-not (Test-Path $DbPath)) {
    Write-Host "Base de dados nao encontrada: $DbPath" -ForegroundColor Red
    exit 1
}

Import-Module PSSQLite -ErrorAction Stop

# ---------------------------------------------------------------------------
# 1. Ler base de dados — ultimos N dias
# ---------------------------------------------------------------------------
Write-Host "A ler historico..." -ForegroundColor Cyan
$cutoff = (Get-Date).AddDays(-$DiasHistorico).ToString("yyyy-MM-dd")
$rows   = Invoke-SqliteQuery -DataSource $DbPath `
            -Query "SELECT * FROM historico WHERE Data >= '$cutoff' ORDER BY Data DESC, Categoria, Marca, Produto"

if ($rows.Count -eq 0) {
    Write-Host "Sem dados nos ultimos $DiasHistorico dias." -ForegroundColor Yellow
    exit 0
}

$dates = $rows | Select-Object -ExpandProperty Data -Unique | Sort-Object -Descending

Write-Host ("  " + $rows.Count + " registos | " + $dates.Count + " datas | desde " + ($dates | Select-Object -Last 1)) -ForegroundColor Green

# ---------------------------------------------------------------------------
# 2. Construir estrutura de dados para o JS
#    { prodId: { Marca, Produto, Categoria, URL, history: { "data": {Preco,PVPR,Desconto,Stock} } } }
# ---------------------------------------------------------------------------
Write-Host "A construir estrutura de dados..." -ForegroundColor Cyan

$products = @{}
foreach ($row in $rows) {
    $id = $row.ProdID
    if (-not $id) { continue }

    if (-not $products.ContainsKey($id)) {
        $products[$id] = @{
            ProdID    = $id
            Marca     = $row.Marca
            Produto   = $row.Produto
            Categoria = $row.Categoria
            URL       = $row.URL
            history   = @{}
        }
    }

    $preco   = if ($null -ne $row.Preco)         { [decimal]$row.Preco         } else { $null }
    $pvpr    = if ($null -ne $row.PVPR)          { [decimal]$row.PVPR          } else { $null }
    $desc    = if ($null -ne $row.Desconto_Pct)  { [int]$row.Desconto_Pct      } else { $null }
    $poup    = if ($null -ne $row.Poupanca_Euro) { [decimal]$row.Poupanca_Euro  } else { $null }

    $products[$id].history[$row.Data] = @{
        Preco         = $preco
        PVPR          = $pvpr
        Desconto_Pct  = $desc
        Poupanca_Euro = $poup
        Stock         = $row.Stock
        Destaque      = $row.Destaque
    }
}

# ---------------------------------------------------------------------------
# 3. Serializar para JSON (PowerShell nativo, sem dependencias)
# ---------------------------------------------------------------------------
function To-JsonValue {
    param($val)
    if ($null -eq $val)          { return "null" }
    if ($val -is [bool])         { return $val.ToString().ToLower() }
    if ($val -is [int])                              { return $val.ToString([System.Globalization.CultureInfo]::InvariantCulture) }
    if ($val -is [decimal] -or $val -is [double])  { return $val.ToString("G", [System.Globalization.CultureInfo]::InvariantCulture) }
    $escaped = $val.ToString() -replace '\\','\\' -replace '"','\"' -replace "`n",' ' -replace "`r",''
    return '"' + $escaped + '"'
}

Write-Host "A serializar JSON..." -ForegroundColor Cyan

$jsonProducts = ($products.Values | ForEach-Object {
    $prod = $_
    $histEntries = ($prod.history.Keys | Sort-Object | ForEach-Object {
        $d = $_
        $h = $prod.history[$d]
        '"' + $d + '":{"p":' + (To-JsonValue $h.Preco) +
            ',"v":' + (To-JsonValue $h.PVPR) +
            ',"d":' + (To-JsonValue $h.Desconto_Pct) +
            ',"s":' + (To-JsonValue ($h.Stock -eq "Sem Stock")) +
            ',"t":' + (To-JsonValue $h.Destaque) + '}'
    }) -join ","

    '{"id":' + (To-JsonValue $prod.ProdID) +
    ',"m":' + (To-JsonValue $prod.Marca) +
    ',"n":' + (To-JsonValue $prod.Produto) +
    ',"c":' + (To-JsonValue $prod.Categoria) +
    ',"u":' + (To-JsonValue $prod.URL) +
    ',"h":{' + $histEntries + '}}'
}) -join ","

$jsonDates = ($dates | ForEach-Object { '"' + $_ + '"' }) -join ","
$jsonCats  = ($rows | Select-Object -ExpandProperty Categoria -Unique | Sort-Object | ForEach-Object { '"' + $_ + '"' }) -join ","
$jsonBrands= ($rows | Select-Object -ExpandProperty Marca -Unique | Sort-Object | ForEach-Object {
    $b = $_ -replace '\\','\\' -replace '"','\"'; '"' + $b + '"'
}) -join ","

# ---------------------------------------------------------------------------
# 4. Gerar HTML
# ---------------------------------------------------------------------------
Write-Host "A gerar dashboard HTML..." -ForegroundColor Cyan

$html = @'
<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Wells Intelligence · Monitor de Preços</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --teal:#008eaa;
  --teal-d:#006d85;
  --teal-dd:#005e73;
  --teal-dim:#e0f5f8;
  --teal-mid:#b3e4ed;
  --green:#0d9488;
  --green-dim:#ccfbf1;
  --red:#e11d48;
  --red-dim:#ffe4e6;
  --bg:#f4fafb;
  --surface:#fff;
  --border:#d1edf2;
  --text:#0f2d35;
  --text-2:#4a7280;
  --text-3:#93b8c0;
  --radius:10px;
  --shadow:0 1px 3px rgba(0,71,85,.07),0 1px 2px rgba(0,71,85,.05);
  --shadow-md:0 4px 12px rgba(0,71,85,.1);
}
body{font-family:'Inter',system-ui,-apple-system,sans-serif;background:var(--bg);color:var(--text);font-size:14px;line-height:1.5}
header{background:var(--teal);color:#fff;padding:12px 32px 14px;display:flex;align-items:flex-end;justify-content:space-between;min-height:100px;flex-shrink:0;position:relative;overflow:hidden}
header::before{content:'';position:absolute;right:-60px;top:-60px;width:260px;height:260px;border-radius:50%;background:rgba(255,255,255,.07);pointer-events:none}
header::after{content:'';position:absolute;right:140px;bottom:-70px;width:160px;height:160px;border-radius:50%;background:rgba(255,255,255,.04);pointer-events:none}
.hdr-left{display:flex;align-items:flex-end;gap:10px;position:relative;z-index:1}
.hdr-logo{height:74px;flex-shrink:0;filter:brightness(0) invert(1)}
.hdr-sub{font-size:11px;color:rgba(255,255,255,.7);letter-spacing:.04em;font-weight:400;line-height:1;white-space:nowrap;text-transform:uppercase}
.hdr-right{font-size:11px;color:rgba(255,255,255,.55);text-align:right;line-height:1.5;position:relative;z-index:1}
nav{background:var(--surface);border-bottom:1px solid var(--border);display:flex;padding:0 32px;box-shadow:0 2px 6px rgba(0,71,85,.06)}
nav button{background:none;border:none;color:var(--text-2);padding:15px 18px;cursor:pointer;font-size:13px;font-weight:500;border-bottom:2px solid transparent;transition:all .15s;white-space:nowrap;font-family:inherit;letter-spacing:.01em}
nav button.active{color:var(--teal-d);border-bottom-color:var(--teal);font-weight:600}
nav button:hover:not(.active){color:var(--text);background:var(--teal-dim)}
.tab{display:none;padding:24px 32px}
.tab.active{display:block}
.filters{display:flex;flex-wrap:wrap;gap:14px;margin-bottom:22px;background:var(--surface);padding:18px 22px;border-radius:var(--radius);box-shadow:var(--shadow);border:1px solid var(--border)}
.filters label{font-size:11px;color:var(--text-2);font-weight:600;text-transform:uppercase;letter-spacing:.05em;display:flex;flex-direction:column;gap:6px}
.filters select,.filters input[type=text],.filters input[type=date]{padding:8px 11px;border:1px solid var(--border);border-radius:7px;font-size:13px;min-width:140px;background:var(--surface);color:var(--text);font-family:inherit;outline:none;transition:border-color .15s,box-shadow .15s}
.filters select:focus,.filters input:focus{border-color:var(--teal);box-shadow:0 0 0 3px rgba(0,142,170,.12)}
.filters input[type=text]{min-width:200px}
.filters input[type=checkbox]{width:15px;height:15px;margin-top:3px;accent-color:var(--teal)}
.btn{background:var(--teal-d);color:#fff;border:none;padding:8px 18px;border-radius:7px;cursor:pointer;font-size:13px;font-weight:500;font-family:inherit;transition:background .15s;letter-spacing:.01em}
.btn:hover{background:var(--teal-dd)}
.btn-compare{background:var(--teal)}.btn-compare:hover{background:var(--teal-d)}
table{width:100%;border-collapse:collapse;background:var(--surface);box-shadow:var(--shadow);border-radius:var(--radius);overflow:hidden}
th{background:var(--teal-dim);color:var(--teal-d);padding:11px 14px;text-align:left;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;white-space:nowrap;cursor:pointer;user-select:none;border-bottom:1px solid var(--teal-mid)}
th:hover{background:#c8eef5;color:var(--teal-dd)}
th.sorted-asc::after{content:" ↑";opacity:.6}
th.sorted-desc::after{content:" ↓";opacity:.6}
td{padding:12px 14px;border-bottom:1px solid var(--border);font-size:13px;line-height:1.45;vertical-align:middle}
tr:last-child td{border-bottom:none}
tbody tr:hover td{background:#f0fafb}
td a{color:var(--teal-d);text-decoration:none;font-weight:500}
td a:hover{color:var(--teal);text-decoration:underline}
.badge{display:inline-flex;align-items:center;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600;letter-spacing:.01em}
.badge-desc{background:var(--red-dim);color:var(--red)}
.badge-bs{background:var(--teal-dim);color:var(--teal-d)}
.badge-eo{background:var(--teal-dim);color:var(--teal-d)}
.badge-novo{background:var(--green-dim);color:var(--green)}
.stock-ok{color:var(--green);font-weight:600}
.stock-no{color:var(--red);font-weight:600}
.price-up{color:var(--red);font-weight:600}
.price-dn{color:var(--green);font-weight:600}
.price-eq{color:var(--text-3)}
.summary-cards{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:22px}
.card{background:var(--surface);border-radius:var(--radius);padding:20px 22px;box-shadow:var(--shadow);border:1px solid var(--border);min-width:140px;text-align:center;flex:1;border-top:3px solid var(--teal);transition:box-shadow .15s,transform .15s}
.card:hover{box-shadow:var(--shadow-md);transform:translateY(-1px)}
.card .val{font-size:30px;font-weight:800;color:var(--teal-d);letter-spacing:-.03em;line-height:1.1}
.card .lbl{font-size:10px;color:var(--text-2);margin-top:7px;font-weight:600;text-transform:uppercase;letter-spacing:.07em}
.chart-wrap{background:var(--surface);border-radius:var(--radius);padding:22px;box-shadow:var(--shadow);border:1px solid var(--border);margin-top:18px}
.chart-wrap canvas{max-height:320px}
.compare-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
.info{color:var(--text-3);font-style:italic;padding:56px;text-align:center;font-size:14px}
.insight-box{background:var(--surface);border-radius:var(--radius);padding:24px;box-shadow:var(--shadow);border:1px solid var(--border);margin-bottom:24px}
.insight-box h2{font-size:16px;font-weight:700;color:var(--teal-d);margin:0 0 16px;letter-spacing:-.01em}
.insight-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:14px;margin-top:6px}
.insight-card{border:1px solid var(--border);border-left:3px solid var(--teal);border-radius:8px;padding:18px 20px;background:var(--surface)}
.insight-card h3{font-size:11px;color:var(--teal-d);margin:0 0 12px;text-transform:uppercase;letter-spacing:.07em;font-weight:700}
.insight-card ul{padding-left:18px;margin:0}
.insight-card li{font-size:13px;line-height:1.7;margin-bottom:10px;padding-left:2px;color:var(--text)}
.insight-card li:last-child{margin-bottom:0}
.insight-full{grid-column:1/-1}
.copy-source{position:absolute;left:-9999px;top:auto;width:1px;height:1px;overflow:hidden}
.section-title{font-size:11px;font-weight:700;color:var(--teal-d);margin:30px 0 12px;text-transform:uppercase;letter-spacing:.08em;display:flex;align-items:center;gap:10px}
.section-title::after{content:'';flex:1;height:1px;background:var(--border)}
.note{font-size:12px;color:var(--text-2);margin-bottom:16px;line-height:1.65;background:var(--teal-dim);padding:10px 14px;border-radius:7px;border-left:3px solid var(--teal)}
#tbl-count,#tbl-count-cmp{font-size:12px;color:var(--text-2);margin-bottom:12px;font-weight:500}
.pagination{display:flex;gap:4px;margin-top:16px;align-items:center;flex-wrap:wrap}
.pagination button{padding:5px 11px;border:1px solid var(--border);background:var(--surface);border-radius:6px;cursor:pointer;font-size:12px;color:var(--text-2);font-family:inherit;transition:all .15s}
.pagination button.active{background:var(--teal-d);color:#fff;border-color:var(--teal-d)}
.pagination button:hover:not(.active){background:var(--teal-dim)}
.evo-cards-grid{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}
.evo-card{display:flex;align-items:center;gap:10px;padding:10px 14px;background:#0f2d35;border:2px solid transparent;border-radius:8px;cursor:pointer;transition:all .18s;min-width:210px;max-width:300px;font-size:13px}
.evo-card:hover{background:#1a3f4a}
.evo-card.selected{border-color:var(--ec,var(--teal));background:#092028}
.evo-dot{width:10px;height:10px;border-radius:50%;background:#2a5563;flex-shrink:0;transition:background .18s}
.evo-card.selected .evo-dot{background:var(--ec,var(--teal))}
.evo-info{flex:1;min-width:0}
.evo-brand{font-size:11px;color:#6fa8b5;display:block}
.evo-name{font-size:12px;color:#c8e8ed;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;display:block}
.evo-price{font-size:13px;font-weight:700;color:#4dd4e8;white-space:nowrap;text-align:right}
.evo-sel-bar{display:flex;align-items:center;gap:12px;padding:10px 14px;background:#0f2d35;border-radius:8px;margin-bottom:12px;font-size:13px;color:#6fa8b5}
.evo-sel-bar span{flex:1}
.brand-grid{display:flex;gap:14px;align-items:flex-start;padding-bottom:12px}
#marcas-grid{overflow-x:auto}
.brand-arrows{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.brand-arrow-btn{background:var(--teal-d);color:#fff;border:none;border-radius:6px;width:34px;height:34px;font-size:20px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:background .15s;user-select:none}
.brand-arrow-btn:hover{background:var(--teal-dd)}
#marcas-range{width:100%;accent-color:var(--teal);cursor:pointer;margin-bottom:8px;display:block}
.brand-col{min-width:190px;max-width:230px;flex-shrink:0;background:var(--surface);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden;border:1px solid var(--border)}
.brand-hdr{background:var(--teal-d);color:#fff;padding:11px 14px;font-weight:600;font-size:13px;display:flex;justify-content:space-between;align-items:center}
.brand-hdr .brand-count{font-size:11px;opacity:.6;font-weight:400}
.brand-prod{padding:10px 14px;border-bottom:1px solid var(--border);font-size:12px}
.brand-prod:last-child{border-bottom:none}
.brand-prod:hover{background:var(--teal-dim)}
.brand-prod-name{display:block;color:var(--teal-d);font-weight:600;margin-bottom:4px;line-height:1.3;text-decoration:none}
.brand-prod-name:hover{color:var(--teal);text-decoration:underline}
.brand-prod-row{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.brand-price{color:var(--teal-d);font-weight:700}
.brand-price-old{color:var(--text-3);text-decoration:line-through;font-size:11px}
</style>
</head>
<body>

<header>
  <div class="hdr-left">
    <img class="hdr-logo" src="assets/Logo.png" alt="Wells Scraper">
    <div class="hdr-sub">Chupetas · Biberões · Bombas Tira-Leite</div>
  </div>
  <div class="hdr-right" id="hdr-meta"></div>
</header>

<nav>
  <button class="active" onclick="showTab('mensal')">Relatório Mensal</button>
  <button onclick="showTab('produtos')">Produtos</button>
  <button onclick="showTab('marcas')">Por Marca</button>
  <button onclick="showTab('evolucao')">Evolução de Preços</button>
  <button onclick="showTab('comparar')">Comparar Datas</button>
</nav>

<!-- ===== TAB: PRODUTOS ===== -->
<div id="tab-produtos" class="tab">
  <div class="filters">
    <label>Data
      <select id="f-date" onchange="renderProducts()"></select>
    </label>
    <label>Categoria
      <select id="f-cat" onchange="renderProducts()"><option value="">Todas</option></select>
    </label>
    <label>Marca
      <select id="f-brand" onchange="renderProducts()"><option value="">Todas</option></select>
    </label>
    <label>Pesquisa
      <input type="text" id="f-search" placeholder="Nome do produto..." oninput="renderProducts()">
    </label>
    <label style="flex-direction:row;align-items:center;gap:8px;padding-top:18px">
      <input type="checkbox" id="f-disc" onchange="renderProducts()"> Só com desconto
    </label>
    <label style="flex-direction:row;align-items:center;gap:8px;padding-top:18px">
      <input type="checkbox" id="f-stock" onchange="renderProducts()"> Só sem stock
    </label>
  </div>
  <div class="summary-cards" id="cards-produtos"></div>
  <div id="tbl-count"></div>
  <table id="tbl-produtos">
    <thead>
      <tr>
        <th onclick="sortProducts(0)" id="ph0">Categoria</th>
        <th onclick="sortProducts(1)" id="ph1">Marca</th>
        <th onclick="sortProducts(2)" id="ph2">Produto</th>
        <th onclick="sortProducts(3)" id="ph3">Preço</th>
        <th onclick="sortProducts(4)" id="ph4">PVPR</th>
        <th onclick="sortProducts(5)" id="ph5">Desconto</th>
        <th onclick="sortProducts(6)" id="ph6">Poupança</th>
        <th onclick="sortProducts(7)" id="ph7">Stock</th>
      </tr>
    </thead>
    <tbody id="tbody-produtos"></tbody>
  </table>
  <div class="pagination" id="pag-produtos"></div>
</div>

<!-- ===== TAB: COMPARAR ===== -->
<div id="tab-comparar" class="tab">
  <div class="filters">
    <label>Data A (anterior)
      <select id="cmp-dateA"></select>
    </label>
    <label>Data B (mais recente)
      <select id="cmp-dateB"></select>
    </label>
    <label>Categoria
      <select id="cmp-cat"><option value="">Todas</option></select>
    </label>
    <label>Mostrar
      <select id="cmp-filter">
        <option value="all">Todos os produtos</option>
        <option value="changed">Só com alteração de preço</option>
        <option value="down">Preço desceu</option>
        <option value="up">Preço subiu</option>
        <option value="new">Novos em B</option>
        <option value="removed">Removidos em A</option>
      </select>
    </label>
    <label style="padding-top:18px">
      <button class="btn btn-compare" onclick="renderCompare()">Comparar</button>
    </label>
    <label style="padding-top:18px">
      <button class="btn" id="btn-cmp-excel" onclick="downloadCompareExcel()" style="display:none">&#11015; Excel</button>
    </label>
  </div>
  <div class="summary-cards" id="cards-compare"></div>
  <div id="tbl-count-cmp"></div>
  <table id="tbl-compare">
    <thead>
      <tr>
        <th onclick="sortCompare(0)" id="ch0">Categoria</th>
        <th onclick="sortCompare(1)" id="ch1">Marca</th>
        <th onclick="sortCompare(2)" id="ch2">Produto</th>
        <th onclick="sortCompare(3)" id="ch3">Preço A</th>
        <th onclick="sortCompare(4)" id="ch4">Preço B</th>
        <th onclick="sortCompare(5)" id="ch5">Diferença €</th>
        <th onclick="sortCompare(6)" id="ch6">Variação %</th>
        <th onclick="sortCompare(7)" id="ch7">Desconto B</th>
        <th onclick="sortCompare(8)" id="ch8">Stock B</th>
      </tr>
    </thead>
    <tbody id="tbody-compare"></tbody>
  </table>
  <div class="pagination" id="pag-compare"></div>
</div>

<!-- ===== TAB: EVOLUCAO ===== -->
<div id="tab-evolucao" class="tab">
  <div class="filters">
    <label>Pesquisa produto
      <input type="text" id="evo-search" placeholder="Nome ou marca..." oninput="searchProducts()">
    </label>
    <label>Categoria
      <select id="evo-cat" onchange="searchProducts()"><option value="">Todas</option></select>
    </label>
  </div>
  <div id="evo-results" style="margin-bottom:16px"></div>
  <div id="evo-export" style="display:none" class="filters">
    <label>Data início
      <input type="date" id="evo-date-start">
    </label>
    <label>Data fim
      <input type="date" id="evo-date-end">
    </label>
    <button class="btn" onclick="downloadExcel()" style="align-self:flex-end">&#11015; Download Excel</button>
  </div>
  <div id="evo-chart-area" style="display:none">
    <div class="summary-cards" id="cards-evo"></div>
    <div class="chart-wrap">
      <canvas id="evoChart"></canvas>
    </div>
    <br>
    <table id="tbl-evo">
      <thead>
        <tr><th>Data</th><th>Preço</th><th>PVPR</th><th>Desconto</th><th>Poupança</th><th>Stock</th></tr>
      </thead>
      <tbody id="tbody-evo"></tbody>
    </table>
  </div>
</div>

<!-- ===== TAB: POR MARCA ===== -->
<div id="tab-marcas" class="tab">
  <div class="filters">
    <label>Categoria
      <select id="marca-cat" onchange="renderMarcas()"><option value="">Todas</option></select>
    </label>
    <label>Ordenar produtos por
      <select id="marca-sort" onchange="renderMarcas()">
        <option value="name">Nome</option>
        <option value="price_asc">Preço ↑</option>
        <option value="price_desc">Preço ↓</option>
        <option value="disc">Desconto</option>
      </select>
    </label>
  </div>
  <div class="brand-arrows">
    <button class="brand-arrow-btn" onclick="scrollMarcas(-1)" title="Anterior">&#8249;</button>
    <button class="brand-arrow-btn" onclick="scrollMarcas(1)" title="Próximo">&#8250;</button>
  </div>
  <input type="range" id="marcas-range" min="0" max="10000" value="0">
  <div id="marcas-grid"></div>
</div>

<!-- ===== TAB: RELATORIO MENSAL ===== -->
<div id="tab-mensal" class="tab active">
  <div class="filters">
    <label>Mês
      <select id="mes-period" onchange="renderMonthlyReport()"></select>
    </label>
    <label>Categoria
      <select id="mes-cat" onchange="renderMonthlyReport()"><option value="">Todas</option></select>
    </label>
    <label>Marca
      <select id="mes-brand" onchange="renderMonthlyReport()"><option value="">Todas</option></select>
    </label>
    <label style="padding-top:18px">
      <button class="btn" onclick="copyMonthlyText()">Copiar insights</button>
    </label>
    <label style="padding-top:18px">
      <button class="btn" onclick="downloadMonthlyExcel()">Excel mensal</button>
    </label>
  </div>

  <div class="summary-cards" id="cards-mensal"></div>

  <div class="insight-box">
    <h2>Insights do mês</h2>
    <div id="monthly-insights"></div>
    <textarea id="monthly-copy" class="copy-source" aria-hidden="true" tabindex="-1"></textarea>
  </div>

  <h2 class="section-title">Novidades do mês</h2>
  <p class="note">Produtos nunca vistos antes no histórico. Esta leitura reflete o site Wells e deve ser cruzada com sell-out e conhecimento comercial.</p>
  <table>
    <thead><tr><th>Categoria</th><th>Marca</th><th>Produto</th><th>Primeiro dia</th><th>Preço</th><th>Desconto</th><th>Stock</th></tr></thead>
    <tbody id="tbody-month-new"></tbody>
  </table>

  <h2 class="section-title">Promoções do mês</h2>
  <table>
    <thead><tr><th>Categoria</th><th>Marca</th><th>Produto</th><th>Dias promo</th><th>Máx desconto</th><th>Preço mínimo</th><th>1ª promo observada</th></tr></thead>
    <tbody id="tbody-month-promos"></tbody>
  </table>

  <h2 class="section-title">Intensidade promocional por categoria</h2>
  <table>
    <thead><tr><th>Categoria</th><th>Refs ativas</th><th>Novidades</th><th>Refs em promo</th><th>Cobertura promo</th><th>Desconto médio</th><th>Máx desconto</th></tr></thead>
    <tbody id="tbody-month-categories"></tbody>
  </table>

  <h2 class="section-title">Intensidade promocional por marca</h2>
  <table>
    <thead><tr><th>Marca</th><th>Refs ativas</th><th>Novidades</th><th>Refs em promo</th><th>Cobertura promo</th><th>Desconto médio</th><th>Máx desconto</th></tr></thead>
    <tbody id="tbody-month-brands"></tbody>
  </table>

  <h2 class="section-title">Presença online: desapareceram no fim do mês</h2>
  <p class="note">Não equivale automaticamente a saída de loja física; é um sinal de presença online a investigar quando for relevante.</p>
  <table>
    <thead><tr><th>Categoria</th><th>Marca</th><th>Produto</th><th>Último dia visto</th><th>Último preço</th><th>Último desconto</th></tr></thead>
    <tbody id="tbody-month-removed"></tbody>
  </table>
</div>

<script>
// ============================================================
// DADOS EMBUTIDOS
// ============================================================
const DATES  = [DATES_JSON];
const CATS   = [CATS_JSON];
const BRANDS = [BRANDS_JSON];
const ALL_PRODUCTS = [PRODUCTS_JSON];
const GENERATED = "TIMESTAMP_VAL";

// ============================================================
// ESTADO
// ============================================================
const PAGE_SIZE = 50;
let sortState = {};
let prodSort = { col: null, dir: 'asc' };
let cmpSort  = { col: null, dir: 'asc' };
let currentPage = { produtos: 1, compare: 1 };
let evoChart = null;
let monthlyData = null;

// ============================================================
// TAB: POR MARCA
// ============================================================
function renderMarcas() {
  const cat   = document.getElementById('marca-cat').value;
  const sortt = document.getElementById('marca-sort').value;
  const grid  = document.getElementById('marcas-grid');

  // Filtrar produtos pela categoria seleccionada
  const prods = ALL_PRODUCTS.filter(p => !cat || p.c === cat);

  // Agrupar por marca
  const byBrand = {};
  prods.forEach(p => {
    if(!byBrand[p.m]) byBrand[p.m] = [];
    // Obter dados mais recentes
    const dates = Object.keys(p.h).sort();
    const last  = dates.length ? p.h[dates[dates.length-1]] : null;
    byBrand[p.m].push({ id:p.id, name:p.n, url:p.u, price:last?last.p:null, pvpr:last?last.v:null, disc:last?last.d:null, stock:last?last.s:false });
  });

  if(!Object.keys(byBrand).length) {
    grid.innerHTML = '<p class="info">Nenhum produto encontrado.</p>';
    return;
  }

  // Ordenar produtos dentro de cada marca
  const sortFn = {
    name:       (a,b) => (a.name||'').localeCompare(b.name||'','pt'),
    price_asc:  (a,b) => (a.price||999)-(b.price||999),
    price_desc: (a,b) => (b.price||0)-(a.price||0),
    disc:       (a,b) => (b.disc||0)-(a.disc||0)
  }[sortt] || ((a,b) => 0);

  // Ordenar marcas por nº de produtos (mais produtos primeiro)
  const brands = Object.keys(byBrand).sort((a,b) => byBrand[b].length - byBrand[a].length);

  grid.innerHTML = '<div class="brand-grid">' + brands.map(brand => {
    const items = [...byBrand[brand]].sort(sortFn);
    const prods = items.map(p => {
      const priceHtml = p.price != null ? '<span class="brand-price">'+fmt(p.price)+'</span>' : '';
      const pvprHtml  = p.pvpr && p.pvpr !== p.price ? '<span class="brand-price-old">'+fmt(p.pvpr)+'</span>' : '';
      const discHtml  = p.disc ? '<span class="badge badge-desc">-'+p.disc+'%</span>' : '';
      const stockHtml = p.stock ? '<span class="stock-no" title="Sem Stock">&#9888;</span>' : '<span class="stock-ok" title="Disponível">&#10003;</span>';
      return '<div class="brand-prod">'
        +'<a class="brand-prod-name" href="'+esc(p.url||'')+'" target="_blank" title="'+esc(p.name)+'">'+esc(p.name.substring(0,42))+'</a>'
        +'<div class="brand-prod-row">'+priceHtml+pvprHtml+discHtml+stockHtml+'</div>'
        +'</div>';
    }).join('');
    return '<div class="brand-col">'
      +'<div class="brand-hdr">'+esc(brand)+'<span class="brand-count">'+items.length+' produto'+(items.length!==1?'s':'')+'</span></div>'
      +prods
      +'</div>';
  }).join('') + '</div>';
  _marcasRangeSync();
}

function scrollMarcas(dir) {
  const grid = document.getElementById('marcas-grid');
  if (!grid) return;
  grid.scrollLeft += dir * (Math.round(window.innerWidth * 0.8) || 800);
}

function _marcasRangeSync() {
  // Actualiza o max do range para corresponder ao scroll real após render
  const grid  = document.getElementById('marcas-grid');
  const range = document.getElementById('marcas-range');
  if (!grid || !range) return;
  const maxScroll = grid.scrollWidth - grid.clientWidth;
  range.max   = maxScroll > 0 ? maxScroll : 1;
  range.value = grid.scrollLeft;
}

// ============================================================
// INIT
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('hdr-meta').textContent =
    'Dados de ' + (DATES[DATES.length-1]||'') + ' a ' + (DATES[0]||'') +
    ' · ' + ALL_PRODUCTS.length + ' produtos · Atualizado ' + GENERATED;

  const fDate = document.getElementById('f-date');
  DATES.forEach(d => { const o=document.createElement('option'); o.value=d; o.textContent=d; fDate.appendChild(o); });

  const cmpA = document.getElementById('cmp-dateA');
  const cmpB = document.getElementById('cmp-dateB');
  DATES.forEach((d,i) => {
    const oA=document.createElement('option'); oA.value=d; oA.textContent=d; cmpA.appendChild(oA);
    const oB=document.createElement('option'); oB.value=d; oB.textContent=d; cmpB.appendChild(oB);
    if(i===1) oA.selected=true;
    if(i===0) oB.selected=true;
  });

  ['f-cat','cmp-cat','evo-cat','marca-cat','mes-cat'].forEach(id => {
    const sel = document.getElementById(id);
    CATS.forEach(c => { const o=document.createElement('option'); o.value=c; o.textContent=c; sel.appendChild(o); });
  });

  const fBrand = document.getElementById('f-brand');
  const mesBrand = document.getElementById('mes-brand');
  BRANDS.forEach(b => {
    const o=document.createElement('option'); o.value=b; o.textContent=b; fBrand.appendChild(o);
    const m=document.createElement('option'); m.value=b; m.textContent=b; mesBrand.appendChild(m);
  });

  const mesPeriod = document.getElementById('mes-period');
  const months = [...new Set(DATES.map(d => d.slice(0,7)))].sort().reverse();
  months.forEach(m => { const o=document.createElement('option'); o.value=m; o.textContent=monthLabel(m); mesPeriod.appendChild(o); });

  renderProducts();
  searchProducts();
  renderMonthlyReport();

  // Range slider <-> grid (sem loop: .value= não dispara 'input')
  const _grid  = document.getElementById('marcas-grid');
  const _range = document.getElementById('marcas-range');
  if (_grid && _range) {
    _range.addEventListener('input', () => { _grid.scrollLeft = +_range.value; });
    _grid.addEventListener('scroll', () => { _range.value = _grid.scrollLeft; }, { passive:true });
  }
});

// ============================================================
// TAB NAVIGATION
// ============================================================
function showTab(name) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('nav button').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
  event.target.classList.add('active');
  if (name === 'marcas') renderMarcas();
  if (name === 'comparar') renderCompare();
  if (name === 'evolucao' && !document.getElementById('evo-search').value) {
    document.getElementById('evo-search').value = 'Chicco';
    searchProducts();
  }
}

// ============================================================
// TAB: PRODUTOS
// ============================================================
function sortProducts(col) {
  if (prodSort.col === col) {
    prodSort.dir = prodSort.dir === 'asc' ? 'desc' : 'asc';
  } else {
    prodSort.col = col;
    prodSort.dir = 'asc';
  }
  // Atualizar indicadores visuais nos headers
  for(let i=0; i<=7; i++) {
    const th = document.getElementById('ph'+i);
    if(!th) continue;
    th.classList.remove('sorted-asc','sorted-desc');
    if(i === prodSort.col) th.classList.add(prodSort.dir==='asc' ? 'sorted-asc' : 'sorted-desc');
  }
  currentPage.produtos = 1;
  _renderProductsPage();
}

function renderProducts() {
  currentPage.produtos = 1;
  _renderProductsPage();
}

function _isNovo(p) {
  const dates = Object.keys(p.h).sort();
  if (!dates.length) return false;
  const firstMs  = new Date(dates[0]).getTime();
  const latestMs = new Date(DATES[0]).getTime();
  return (latestMs - firstMs) <= 14 * 864e5;
}

function _renderProductsPage() {
  const date   = document.getElementById('f-date').value;
  const cat    = document.getElementById('f-cat').value;
  const brand  = document.getElementById('f-brand').value;
  const search = document.getElementById('f-search').value.toLowerCase().trim();
  const onlyDisc  = document.getElementById('f-disc').checked;
  const onlyStock = document.getElementById('f-stock').checked;

  const filtered = ALL_PRODUCTS.filter(p => {
    const h = p.h[date]; if(!h) return false;
    if(cat   && p.c !== cat)   return false;
    if(brand && p.m !== brand) return false;
    if(search && !p.n.toLowerCase().includes(search) && !p.m.toLowerCase().includes(search)) return false;
    if(onlyDisc  && !h.d) return false;
    if(onlyStock && !h.s) return false;
    return true;
  });

  // Ordenar todos os dados antes de paginar
  if (prodSort.col !== null) {
    const col = prodSort.col;
    const asc = prodSort.dir === 'asc';
    filtered.sort((a, b) => {
      const ha = a.h[date], hb = b.h[date];
      let va, vb;
      if      (col===0) { va=a.c;           vb=b.c; }
      else if (col===1) { va=a.m;           vb=b.m; }
      else if (col===2) { va=a.n;           vb=b.n; }
      else if (col===3) { va=ha?.p||0;      vb=hb?.p||0; }
      else if (col===4) { va=ha?.v||0;      vb=hb?.v||0; }
      else if (col===5) { va=ha?.d||0;      vb=hb?.d||0; }
      else if (col===6) { va=(ha?.v&&ha?.p)?ha.v-ha.p:0; vb=(hb?.v&&hb?.p)?hb.v-hb.p:0; }
      else if (col===7) { va=ha?.s?1:0;     vb=hb?.s?1:0; }
      if (typeof va==='string') return asc ? va.localeCompare(vb,'pt') : vb.localeCompare(va,'pt');
      return asc ? va-vb : vb-va;
    });
  }

  // Novidades sempre no topo (independente da ordenação)
  const novos = filtered.filter(p => _isNovo(p));
  const resto = filtered.filter(p => !_isNovo(p));
  const sorted = [...novos, ...resto];

  // Summary cards
  const total    = filtered.length;
  const withDisc = filtered.filter(p => p.h[date] && p.h[date].d).length;
  const noStock  = filtered.filter(p => p.h[date] && p.h[date].s).length;
  const maxDisc  = filtered.reduce((m,p) => { const d=p.h[date]?.d||0; return d>m?d:m; }, 0);
  const minPrice = filtered.reduce((m,p) => { const pr=p.h[date]?.p||9999; return pr<m?pr:m; }, 9999);

  document.getElementById('cards-produtos').innerHTML =
    card(total,        'Produtos') +
    card(withDisc,     'Com desconto') +
    card(noStock,      'Sem stock') +
    card(maxDisc+'%',  'Maior desconto') +
    (minPrice<9999 ? card('€'+minPrice.toFixed(2), 'Preço mais baixo') : '');

  // Pagination
  const pages = Math.ceil(sorted.length / PAGE_SIZE);
  const p = currentPage.produtos;
  const slice = sorted.slice((p-1)*PAGE_SIZE, p*PAGE_SIZE);

  document.getElementById('tbl-count').textContent =
    sorted.length + ' produtos encontrados' + (novos.length ? ' · ' + novos.length + ' novidade' + (novos.length>1?'s':'') : '') + (pages>1 ? ' (página '+p+' de '+pages+')' : '');

  const tbody = document.getElementById('tbody-produtos');
  tbody.innerHTML = slice.map(prod => {
    const h = prod.h[date];
    const novoBadge = _isNovo(prod) ? ' <span class="badge badge-novo">NOVO</span>' : '';
    const dCell = h.d ? '<span class="badge badge-desc">-'+h.d+'%</span>' : '-';
    const sCell = h.s ? '<span class="stock-no">Sem Stock</span>' : '<span class="stock-ok">Disponível</span>';
    return '<tr>' +
      '<td>'+esc(prod.c)+'</td>' +
      '<td>'+esc(prod.m)+'</td>' +
      '<td><a href="'+esc(prod.u)+'" target="_blank">'+esc(prod.n)+'</a>'+novoBadge+'</td>' +
      '<td>'+fmt(h.p)+'</td>' +
      '<td>'+(h.v?fmt(h.v):'-')+'</td>' +
      '<td>'+dCell+'</td>' +
      '<td>'+(h.v&&h.p?'<span style="color:#1e8449">€'+(h.v-h.p).toFixed(2)+'</span>':'-')+'</td>' +
      '<td>'+sCell+'</td>' +
      '</tr>';
  }).join('');

  renderPagination('pag-produtos', pages, p, (n)=>{ currentPage.produtos=n; _renderProductsPage(); });
}

// ============================================================
// TAB: COMPARAR
// ============================================================
function renderCompare() {
  cmpSort = { col: null, dir: 'asc' };
  currentPage.compare = 1;
  _renderComparePage();
}

function sortCompare(col) {
  if (cmpSort.col === col) {
    cmpSort.dir = cmpSort.dir === 'asc' ? 'desc' : 'asc';
  } else {
    cmpSort.col = col;
    cmpSort.dir = 'asc';
  }
  currentPage.compare = 1;
  _renderComparePage();
}

function _renderComparePage() {
  const dateA  = document.getElementById('cmp-dateA').value;
  const dateB  = document.getElementById('cmp-dateB').value;
  const cat    = document.getElementById('cmp-cat').value;
  const filter = document.getElementById('cmp-filter').value;

  if(dateA === dateB) {
    document.getElementById('tbody-compare').innerHTML = '<tr><td colspan="9" class="info">Seleciona duas datas diferentes.</td></tr>';
    return;
  }

  const allIds = new Set([
    ...ALL_PRODUCTS.filter(p=>p.h[dateA]).map(p=>p.id),
    ...ALL_PRODUCTS.filter(p=>p.h[dateB]).map(p=>p.id)
  ]);

  let rows = [];
  allIds.forEach(id => {
    const prod = ALL_PRODUCTS.find(p=>p.id===id);
    if(!prod) return;
    if(cat && prod.c !== cat) return;

    const hA = prod.h[dateA];
    const hB = prod.h[dateB];
    const isNew     = !hA && hB;
    const isRemoved = hA && !hB;
    const pA = hA?.p ?? null;
    const pB = hB?.p ?? null;
    const diff = (pA!==null&&pB!==null) ? pB-pA : null;
    const pct  = (pA!==null&&pB!==null&&pA>0) ? ((pB-pA)/pA)*100 : null;

    const type = isNew?'new': isRemoved?'removed': (diff===null?'eq': diff<0?'down': diff>0?'up':'eq');

    if(filter==='changed' && type==='eq') return;
    if(filter==='down'    && type!=='down') return;
    if(filter==='up'      && type!=='up') return;
    if(filter==='new'     && type!=='new') return;
    if(filter==='removed' && type!=='removed') return;

    rows.push({ prod, hA, hB, pA, pB, diff, pct, type });
  });

  // Ordenação: coluna escolhida pelo utilizador, ou por defeito maior variação absoluta
  if (cmpSort.col !== null) {
    const col = cmpSort.col;
    const asc = cmpSort.dir === 'asc';
    rows.sort((a, b) => {
      let va, vb;
      if      (col===0) { va=a.prod.c;       vb=b.prod.c; }
      else if (col===1) { va=a.prod.m;       vb=b.prod.m; }
      else if (col===2) { va=a.prod.n;       vb=b.prod.n; }
      else if (col===3) { va=a.pA??-999;     vb=b.pA??-999; }
      else if (col===4) { va=a.pB??-999;     vb=b.pB??-999; }
      else if (col===5) { va=a.diff??-999;   vb=b.diff??-999; }
      else if (col===6) { va=a.pct??-999;    vb=b.pct??-999; }
      else if (col===7) { va=a.hB?.d||0;     vb=b.hB?.d||0; }
      else if (col===8) { va=a.hB?.s?1:0;    vb=b.hB?.s?1:0; }
      if (typeof va==='string') return asc ? va.localeCompare(vb,'pt') : vb.localeCompare(va,'pt');
      return asc ? va-vb : vb-va;
    });
  } else {
    // Por defeito: novos (catálogo ou só em B) e removidos no topo, depois maior variação absoluta
    const typeOrder = { new:0, removed:1, down:2, up:3, eq:4 };
    rows.sort((a,b) => {
      const aOrder = (a.type==='new' || _isNovo(a.prod)) ? 0 : (typeOrder[a.type]||4);
      const bOrder = (b.type==='new' || _isNovo(b.prod)) ? 0 : (typeOrder[b.type]||4);
      if (aOrder !== bOrder) return aOrder - bOrder;
      return Math.abs(b.diff||0) - Math.abs(a.diff||0);
    });
  }

  // Actualizar indicadores visuais nos headers
  for(let i=0;i<=8;i++){const th=document.getElementById('ch'+i);if(th){th.classList.remove('sorted-asc','sorted-desc');if(i===cmpSort.col)th.classList.add(cmpSort.dir==='asc'?'sorted-asc':'sorted-desc');}}

  // Cards
  const nDown    = rows.filter(r=>r.type==='down').length;
  const nUp      = rows.filter(r=>r.type==='up').length;
  const nNew     = rows.filter(r=>r.type==='new').length;
  const nRemoved = rows.filter(r=>r.type==='removed').length;
  const maxDrop  = rows.filter(r=>r.diff!==null).reduce((m,r)=>r.diff<m?r.diff:m,0);

  document.getElementById('cards-compare').innerHTML =
    card('<span style="color:#1e8449">'+nDown+'</span>', 'Preço desceu') +
    card('<span style="color:#c0392b">'+nUp+'</span>',   'Preço subiu') +
    card(nNew,     'Novos produtos') +
    card(nRemoved, 'Removidos') +
    (maxDrop<0 ? card('<span style="color:#1e8449">€'+maxDrop.toFixed(2)+'</span>', 'Maior descida') : '');

  const pages = Math.ceil(rows.length / PAGE_SIZE);
  const p = currentPage.compare;
  const slice = rows.slice((p-1)*PAGE_SIZE, p*PAGE_SIZE);

  document.getElementById('tbl-count-cmp').textContent =
    rows.length + ' produtos' + (pages>1 ? ' (página '+p+' de '+pages+')' : '');

  const tbody = document.getElementById('tbody-compare');
  tbody.innerHTML = slice.map(r => {
    const { prod, hA, hB, pA, pB, diff, pct, type } = r;
    const cls = type==='down'?'price-dn': type==='up'?'price-up': type==='new'?'price-dn': type==='removed'?'price-up':'price-eq';
    const prA = pA!==null ? fmt(pA) : '<span class="price-up">—</span>';
    const prB = pB!==null ? fmt(pB) : '<span class="price-up">—</span>';
    const diffStr = diff!==null ? '<span class="'+cls+'">'+(diff>=0?'+':'')+diff.toFixed(2)+'€</span>' : '-';
    const pctStr  = pct!==null  ? '<span class="'+cls+'">'+(pct>=0?'+':'')+pct.toFixed(1)+'%</span>'  : '-';
    const dB  = hB?.d ? '<span class="badge badge-desc">-'+hB.d+'%</span>' : '-';
    const stB = hB ? (hB.s?'<span class="stock-no">Sem Stock</span>':'<span class="stock-ok">OK</span>') : '-';
    const isNovoCat = _isNovo(prod);
    const novoBadge     = (type==='new' || isNovoCat) ? ' <span class="badge badge-novo">NOVO</span>'     : '';
    const removidoBadge = type==='removed' ? ' <span class="badge" style="background:#c0392b;color:#fff">REMOVIDO</span>' : '';
    const rowCls = (type==='new'||isNovoCat)?'style="background:#f0fff4"': type==='removed'?'style="background:#fff5f5"':'';
    return '<tr '+rowCls+'>' +
      '<td>'+esc(prod.c)+'</td>' +
      '<td>'+esc(prod.m)+'</td>' +
      '<td><a href="'+esc(prod.u)+'" target="_blank">'+esc(prod.n)+'</a>'+novoBadge+removidoBadge+'</td>' +
      '<td>'+prA+'</td><td>'+prB+'</td>' +
      '<td>'+diffStr+'</td><td>'+pctStr+'</td>' +
      '<td>'+dB+'</td><td>'+stB+'</td>' +
      '</tr>';
  }).join('');

  renderPagination('pag-compare', pages, p, (n)=>{ currentPage.compare=n; _renderComparePage(); });

  // Guardar rows para export e mostrar/esconder botão Excel
  window._cmpRows = rows;
  window._cmpDates = { a: dateA, b: dateB };
  document.getElementById('btn-cmp-excel').style.display = rows.length ? '' : 'none';
}

function downloadCompareExcel() {
  const rows = window._cmpRows;
  if (!rows || !rows.length) return;
  const { a: dateA, b: dateB } = window._cmpDates || {};

  const header = ['Categoria','Marca','Produto','Preço '+dateA,'Preço '+dateB,'Variação €','Variação %','Desconto '+dateB,'Stock '+dateB];
  const data = [header, ...rows.map(r => {
    const { prod, hA, hB, pA, pB, diff, pct, type } = r;
    return [
      prod.c,
      prod.m,
      prod.n,
      pA !== null ? pA : '',
      pB !== null ? pB : '',
      diff !== null ? +diff.toFixed(2) : '',
      pct  !== null ? +pct.toFixed(1)  : (type==='new'?'Novo':'Removido'),
      hB?.d ? '-'+hB.d+'%' : '',
      hB ? (hB.s ? 'Sem Stock' : 'Disponível') : ''
    ];
  })];

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [14,18,45,14,14,12,12,12,14].map(w=>({wch:w}));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Comparação');
  XLSX.writeFile(wb, 'Comparacao_'+dateA+'_vs_'+dateB+'.xlsx');
}

// ============================================================
// TAB: EVOLUCAO — comparacao multipla
// ============================================================
const EVO_COLORS = ['#2874a6','#e74c3c','#27ae60','#f39c12','#8e44ad'];
let evoSelected = []; // array de IDs seleccionados (max 5)

function _evoSearch() {
  const q   = document.getElementById('evo-search').value.toLowerCase().trim();
  const cat = document.getElementById('evo-cat').value;
  return ALL_PRODUCTS.filter(p => {
    if(cat && p.c !== cat) return false;
    return !q || p.n.toLowerCase().includes(q) || p.m.toLowerCase().includes(q);
  }).slice(0, 30);
}

function searchProducts() {
  const div   = document.getElementById('evo-results');
  const q     = document.getElementById('evo-search').value.toLowerCase().trim();
  const cat   = document.getElementById('evo-cat').value;
  if(!q && !cat) { div.innerHTML = '<p class="info">Pesquisa um produto pelo nome ou marca. Podes seleccionar até 5 para comparar no mesmo gráfico.</p>'; return; }
  const found = _evoSearch();
  if(!found.length) { div.innerHTML = '<p class="info">Nenhum produto encontrado.</p>'; return; }
  _renderEvoCards(found);
}

function _renderEvoCards(found) {
  const div = document.getElementById('evo-results');
  const bar = evoSelected.length > 0
    ? '<div class="evo-sel-bar"><span>'+evoSelected.length+' produto'+(evoSelected.length>1?'s':'')+' seleccionado'+(evoSelected.length>1?'s':'')+' (máx. 5)</span><button class="btn" style="font-size:12px;padding:5px 12px" onclick="clearEvoSelection()">Limpar seleção</button></div>'
    : '';
  const cards = found.map(p => {
    const si    = evoSelected.indexOf(p.id);
    const sel   = si >= 0;
    const color = sel ? EVO_COLORS[si % EVO_COLORS.length] : '#555';
    const dates = Object.keys(p.h).sort();
    const last  = dates.length ? p.h[dates[dates.length-1]] : null;
    const price = last ? fmt(last.p) : '-';
    const disc  = last && last.d ? ' <span class="badge badge-desc">-'+last.d+'%</span>' : '';
    return '<div class="evo-card'+(sel?' selected':'')+'" style="--ec:'+color+'" onclick="toggleEvoProduct(\''+p.id+'\')">'
      +'<div class="evo-dot"></div>'
      +'<div class="evo-info"><span class="evo-brand">'+esc(p.m)+'</span><span class="evo-name" title="'+esc(p.n)+'">'+esc(p.n.substring(0,38))+'</span></div>'
      +'<div class="evo-price">'+price+disc+'</div>'
      +'</div>';
  }).join('');
  div.innerHTML = bar + '<div class="evo-cards-grid">'+cards+'</div>';
}

function toggleEvoProduct(id) {
  const idx = evoSelected.indexOf(id);
  if(idx >= 0) {
    evoSelected.splice(idx, 1);
  } else {
    if(evoSelected.length >= 5) return;
    evoSelected.push(id);
  }
  _renderEvoCards(_evoSearch());
  if(evoSelected.length > 0) { updateEvoChart(); }
  else { document.getElementById('evo-chart-area').style.display='none'; document.getElementById('evo-export').style.display='none'; }
}

function clearEvoSelection() {
  evoSelected = [];
  _renderEvoCards(_evoSearch());
  document.getElementById('evo-chart-area').style.display='none';
  document.getElementById('evo-export').style.display='none';
}

function updateEvoChart() {
  const prods = evoSelected.map(id => ALL_PRODUCTS.find(p => p.id===id)).filter(Boolean);
  if(!prods.length) return;
  const single = prods.length === 1;

  // Datas union de todos os produtos seleccionados
  const allDates = [...new Set(prods.flatMap(p => Object.keys(p.h)))].sort();

  // Export panel
  document.getElementById('evo-export').style.display = 'flex';
  document.getElementById('evo-date-start').value = allDates[0]||'';
  document.getElementById('evo-date-end').value   = allDates[allDates.length-1]||'';

  // Cards de estatísticas
  if(single) {
    const p = prods[0]; const dd = Object.keys(p.h).sort();
    const pp = dd.map(d=>p.h[d].p).filter(v=>v!=null);
    const lastP=pp[pp.length-1], firstP=pp[0], minP=Math.min(...pp), maxP=Math.max(...pp);
    const tv = firstP>0 ? ((lastP-firstP)/firstP*100) : 0;
    document.getElementById('cards-evo').innerHTML =
      card(fmt(lastP),'Preço atual')+card(fmt(minP),'Mínimo histórico')+card(fmt(maxP),'Máximo histórico')+
      card((tv>=0?'+':'')+tv.toFixed(1)+'%','Variação total')+card(dd.length,'Dias registados');
  } else {
    document.getElementById('cards-evo').innerHTML = prods.map((p,i) => {
      const col = EVO_COLORS[i%EVO_COLORS.length];
      const dd  = Object.keys(p.h).sort(); const pp = dd.map(d=>p.h[d].p).filter(v=>v!=null);
      const lastP=pp[pp.length-1], minP=Math.min(...pp);
      const tv = pp.length>1 ? ((lastP-pp[0])/pp[0]*100) : 0;
      return '<div class="card" style="border-top:3px solid '+col+';min-width:120px">'
        +'<div class="val" style="font-size:12px;color:'+col+'">'+esc(p.m.substring(0,14))+'</div>'
        +'<div class="val" style="font-size:20px">'+fmt(lastP)+'</div>'
        +'<div class="lbl">mín '+fmt(minP)+' &nbsp;·&nbsp; '+(tv>=0?'+':'')+tv.toFixed(1)+'%</div>'
        +'</div>';
    }).join('');
  }

  // Gráfico
  document.getElementById('evo-chart-area').style.display = 'block';
  if(evoChart) evoChart.destroy();
  const ctx = document.getElementById('evoChart').getContext('2d');
  const datasets = prods.flatMap((p,i) => {
    const col  = EVO_COLORS[i%EVO_COLORS.length];
    const prices = allDates.map(d => p.h[d] ? p.h[d].p : null);
    const ds = [{ label: p.m+' – '+p.n.substring(0,28), data: prices,
      borderColor:col, backgroundColor: single?col.replace('#','rgba(')+'1)'.replace(/rgba\(([0-9a-f]{6})1\)/,(_,h)=>'rgba('+parseInt(h.substring(0,2),16)+','+parseInt(h.substring(2,4),16)+','+parseInt(h.substring(4,6),16)+',.1)'):'transparent',
      tension:0.3, fill:single, pointRadius:4, spanGaps:true }];
    if(single) ds.push({ label:'PVPR', data:allDates.map(d=>p.h[d]?(p.h[d].v||null):null),
      borderColor:'#e74c3c', borderDash:[5,5], tension:0.3, fill:false, pointRadius:2, spanGaps:true });
    return ds;
  });
  evoChart = new Chart(ctx, {
    type:'line', data:{labels:allDates, datasets},
    options:{ responsive:true, plugins:{
      title:{display:true, text: single?(prods[0].m+' — '+prods[0].n):'Comparação de preços', font:{size:13}},
      tooltip:{callbacks:{label:c=>c.dataset.label+': €'+(c.parsed.y||0).toFixed(2)}}
    }, scales:{y:{ticks:{callback:v=>'€'+v.toFixed(2)}}}}
  });

  // Tabela de histórico — só para produto único
  document.getElementById('tbl-evo').style.display = single ? '' : 'none';
  if(single) {
    const p = prods[0]; const dd = Object.keys(p.h).sort();
    document.getElementById('tbody-evo').innerHTML = [...dd].reverse().map(d => {
      const h=p.h[d];
      return '<tr><td>'+d+'</td><td>'+fmt(h.p)+'</td><td>'+(h.v?fmt(h.v):'-')+'</td>'
        +'<td>'+(h.d?'<span class="badge badge-desc">-'+h.d+'%</span>':'-')+'</td>'
        +'<td>'+(h.v&&h.p?'€'+(h.v-h.p).toFixed(2):'-')+'</td>'
        +'<td>'+(h.s?'<span class="stock-no">Sem Stock</span>':'<span class="stock-ok">OK</span>')+'</td></tr>';
    }).join('');
  }
}

// Manter compatibilidade com chamadas antigas
function selectProduct(id) { toggleEvoProduct(id); }

// ============================================================
// TAB: RELATORIO MENSAL
// ============================================================
function monthLabel(m) {
  const nomes = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
  const [y, mo] = m.split('-');
  return nomes[Number(mo)-1] + ' de ' + y;
}

function getMonthDates(month) {
  return DATES.filter(d => d.startsWith(month)).sort();
}

function firstSeenDate(prod) {
  const ds = Object.keys(prod.h).sort();
  return ds.length ? ds[0] : null;
}

function lastRecordInDates(prod, dates) {
  for(let i=dates.length-1; i>=0; i--) {
    const d = dates[i];
    if(prod.h[d]) return { date:d, h:prod.h[d] };
  }
  return null;
}

function hasRecordInDates(prod, dates) {
  return dates.some(d => !!prod.h[d]);
}

function productPassesMonthlyFilters(prod, cat, brand) {
  if(cat && prod.c !== cat) return false;
  if(brand && prod.m !== brand) return false;
  return true;
}

function renderMonthlyReport() {
  const month = document.getElementById('mes-period').value;
  const cat   = document.getElementById('mes-cat').value;
  const brand = document.getElementById('mes-brand').value;
  const dates = getMonthDates(month);

  if(!month || !dates.length) {
    document.getElementById('cards-mensal').innerHTML = '';
    document.getElementById('monthly-copy').value = '';
    document.getElementById('monthly-insights').innerHTML = '';
    return;
  }

  const firstDate = dates[0];
  const lastDate  = dates[dates.length-1];
  const active = ALL_PRODUCTS
    .filter(p => productPassesMonthlyFilters(p, cat, brand))
    .filter(p => hasRecordInDates(p, dates));

  const newProducts = active
    .filter(p => firstSeenDate(p)?.startsWith(month))
    .map(p => {
      const first = firstSeenDate(p);
      return { prod:p, first:first, h:p.h[first] };
    })
    .sort((a,b) => a.prod.c.localeCompare(b.prod.c,'pt') || a.prod.m.localeCompare(b.prod.m,'pt') || a.first.localeCompare(b.first));

  const promos = active.map(p => {
    const promoDates = dates.filter(d => p.h[d] && p.h[d].d);
    if(!promoDates.length) return null;
    const discounts = promoDates.map(d => Number(p.h[d].d || 0));
    const priceDates = dates.filter(d => p.h[d] && p.h[d].p != null);
    const prices = priceDates.map(d => Number(p.h[d].p));
    const previousPromo = Object.keys(p.h).some(d => d < firstDate && p.h[d].d);
    return {
      prod:p,
      days: promoDates.length,
      maxDisc: Math.max(...discounts),
      avgDisc: discounts.reduce((a,b)=>a+b,0) / discounts.length,
      minPrice: prices.length ? Math.min(...prices) : null,
      firstPromo: promoDates[0],
      newPromo: !previousPromo,
      discountObs: discounts.length,
      discountSum: discounts.reduce((a,b)=>a+b,0)
    };
  }).filter(Boolean).sort((a,b) => b.maxDisc-a.maxDisc || b.days-a.days);

  const removed = active
    .filter(p => !p.h[lastDate])
    .map(p => ({ prod:p, last:lastRecordInDates(p, dates) }))
    .filter(x => x.last)
    .sort((a,b) => a.last.date.localeCompare(b.last.date) || a.prod.m.localeCompare(b.prod.m,'pt'));

  const brandStats = {};
  active.forEach(p => {
    if(!brandStats[p.m]) brandStats[p.m] = { brand:p.m, active:0, newCount:0, promoRefs:0, promoDays:0, discSum:0, discObs:0, maxDisc:0 };
    brandStats[p.m].active++;
  });
  newProducts.forEach(x => { if(brandStats[x.prod.m]) brandStats[x.prod.m].newCount++; });
  promos.forEach(x => {
    const s = brandStats[x.prod.m];
    if(!s) return;
    s.promoRefs++;
    s.promoDays += x.days;
    s.discSum += x.discountSum;
    s.discObs += x.discountObs;
    s.maxDisc = Math.max(s.maxDisc, x.maxDisc);
  });
  const brandRows = Object.values(brandStats).sort((a,b) => b.promoRefs-a.promoRefs || b.newCount-a.newCount || a.brand.localeCompare(b.brand,'pt'));

  const categoryStats = {};
  active.forEach(p => {
    if(!categoryStats[p.c]) categoryStats[p.c] = { category:p.c, active:0, newCount:0, promoRefs:0, promoDays:0, discSum:0, discObs:0, maxDisc:0 };
    categoryStats[p.c].active++;
  });
  newProducts.forEach(x => { if(categoryStats[x.prod.c]) categoryStats[x.prod.c].newCount++; });
  promos.forEach(x => {
    const s = categoryStats[x.prod.c];
    if(!s) return;
    s.promoRefs++;
    s.promoDays += x.days;
    s.discSum += x.discountSum;
    s.discObs += x.discountObs;
    s.maxDisc = Math.max(s.maxDisc, x.maxDisc);
  });
  const categoryRows = Object.values(categoryStats).sort((a,b) => b.promoRefs-a.promoRefs || a.category.localeCompare(b.category,'pt'));

  monthlyData = { month, dates, active, newProducts, promos, removed, brandRows, categoryRows };

  document.getElementById('cards-mensal').innerHTML =
    card(active.length, 'Refs online observadas') +
    card(newProducts.length, 'Novidades') +
    card(promos.length, 'Refs em promoção') +
    card(promos.reduce((m,p)=>Math.max(m,p.maxDisc),0)+'%', 'Maior desconto') +
    card(removed.length, 'Ausentes no fim do período');

  const insights = buildMonthlyInsights(month, dates, newProducts, promos, removed, brandRows, categoryRows, cat, brand);
  document.getElementById('monthly-insights').innerHTML = renderMonthlyInsights(insights);
  document.getElementById('monthly-copy').value = insightsToText(insights);

  renderMonthlyTables(newProducts, promos, brandRows, categoryRows, removed);
}

function renderMonthlyTables(newProducts, promos, brandRows, categoryRows, removed) {
  document.getElementById('tbody-month-new').innerHTML = newProducts.length ? newProducts.map(x =>
    '<tr><td>'+esc(x.prod.c)+'</td><td>'+esc(x.prod.m)+'</td><td><a href="'+esc(x.prod.u)+'" target="_blank">'+esc(x.prod.n)+'</a></td><td>'+x.first+'</td><td>'+fmt(x.h.p)+'</td><td>'+(x.h.d?'-'+x.h.d+'%':'-')+'</td><td>'+(x.h.s?'Sem Stock':'Disponível')+'</td></tr>'
  ).join('') : '<tr><td colspan="7" class="info">Sem novidades no período selecionado.</td></tr>';

  document.getElementById('tbody-month-promos').innerHTML = promos.length ? promos.slice(0,120).map(x =>
    '<tr><td>'+esc(x.prod.c)+'</td><td>'+esc(x.prod.m)+'</td><td><a href="'+esc(x.prod.u)+'" target="_blank">'+esc(x.prod.n)+'</a></td><td>'+x.days+'</td><td><span class="badge badge-desc">-'+x.maxDisc+'%</span></td><td>'+fmt(x.minPrice)+'</td><td>'+(x.newPromo?'Sim':'Não')+'</td></tr>'
  ).join('') : '<tr><td colspan="7" class="info">Sem promoções no período selecionado.</td></tr>';

  const _cov = (x,n) => x.promoRefs && n ? Math.round(x.promoDays/x.promoRefs/n*100)+'%' : '-';
  document.getElementById('tbody-month-categories').innerHTML = categoryRows.length ? categoryRows.map(x =>
    '<tr><td>'+esc(x.category)+'</td><td>'+x.active+'</td><td>'+x.newCount+'</td><td>'+x.promoRefs+'</td><td>'+_cov(x,monthlyData.dates.length)+'</td><td>'+(x.discObs?(x.discSum/x.discObs).toFixed(1)+'%':'-')+'</td><td>'+(x.maxDisc?x.maxDisc+'%':'-')+'</td></tr>'
  ).join('') : '<tr><td colspan="7" class="info">Sem dados por categoria no período selecionado.</td></tr>';

  document.getElementById('tbody-month-brands').innerHTML = brandRows.length ? brandRows.map(x =>
    '<tr><td>'+esc(x.brand)+'</td><td>'+x.active+'</td><td>'+x.newCount+'</td><td>'+x.promoRefs+'</td><td>'+_cov(x,monthlyData.dates.length)+'</td><td>'+(x.discObs?(x.discSum/x.discObs).toFixed(1)+'%':'-')+'</td><td>'+(x.maxDisc?x.maxDisc+'%':'-')+'</td></tr>'
  ).join('') : '<tr><td colspan="7" class="info">Sem dados no período selecionado.</td></tr>';

  document.getElementById('tbody-month-removed').innerHTML = removed.length ? removed.map(x =>
    '<tr><td>'+esc(x.prod.c)+'</td><td>'+esc(x.prod.m)+'</td><td><a href="'+esc(x.prod.u)+'" target="_blank">'+esc(x.prod.n)+'</a></td><td>'+x.last.date+'</td><td>'+fmt(x.last.h.p)+'</td><td>'+(x.last.h.d?'-'+x.last.h.d+'%':'-')+'</td></tr>'
  ).join('') : '<tr><td colspan="6" class="info">Sem ausências no fim do período selecionado.</td></tr>';
}

function buildMonthlyInsights(month, dates, newProducts, promos, removed, brandRows, categoryRows, cat, brand) {
  const periodo = monthLabel(month);
  const scope = [cat || 'todas as categorias', brand ? 'marca '+brand : 'todas as marcas'].join(' / ');
  const insights = [
    {
      title: 'Leitura geral',
      items: [
        'Período observado: ' + periodo + ' (' + scope + ').',
        'Foram observadas ' + monthlyData.active.length + ' referências online no site Wells.',
        'Esta leitura deve ser cruzada com sell-out, informação comercial e disponibilidade em loja física.'
      ]
    },
    {
      title: 'Novidades',
      items: []
    },
    {
      title: 'Promoções',
      items: []
    },
    {
      title: 'Watch outs',
      items: []
    }
  ];

  if(newProducts.length) {
    const byCat = {};
    newProducts.forEach(x => {
      if(!byCat[x.prod.c]) byCat[x.prod.c] = {};
      if(!byCat[x.prod.c][x.prod.m]) byCat[x.prod.c][x.prod.m] = [];
      byCat[x.prod.c][x.prod.m].push(x.prod.n);
    });
    Object.keys(byCat).sort().forEach(c => {
      const brandBits = Object.keys(byCat[c]).sort().map(b => {
        const names = byCat[c][b].slice(0,4).map(shortProductName).join(', ');
        const extra = byCat[c][b].length > 4 ? ' e mais ' + (byCat[c][b].length-4) + ' ref.' : '';
        return b + ' (' + byCat[c][b].length + '): ' + names + extra;
      });
      insights[1].items.push(c + ': ' + brandBits.join('; ') + '.');
    });
  } else {
    insights[1].items.push('Não foram identificadas novas referências no período selecionado.');
  }

  if(promos.length) {
    const topPromos = promos.slice(0,3).map(p => p.prod.m + ' ' + shortProductName(p.prod.n) + ' (-' + p.maxDisc + '%)').join('; ');
    const firstPromos = promos.filter(p => p.newPromo).length;
    const avgPromo = promos.reduce((s,p)=>s+p.avgDisc,0) / promos.length;
    insights[2].items.push('Foram observadas ' + promos.length + ' referências em promoção.');
    insights[2].items.push('Maiores sinais promocionais: ' + topPromos + '.');
    insights[2].items.push(firstPromos + ' referência(s) tiveram a primeira promoção observada no histórico disponível.');
    insights[2].items.push('Desconto médio por referência promocionada: ' + avgPromo.toFixed(1) + '%.');
    if(categoryRows.length) {
      const categorySummary = categoryRows.map(c => {
        const avg = c.discObs ? (c.discSum/c.discObs).toFixed(1) + '%' : '-';
        return c.category + ': ' + c.promoRefs + ' refs em promo, max ' + (c.maxDisc || 0) + '%, desconto médio ' + avg;
      }).join('; ');
      insights[2].items.push('Por categoria: ' + categorySummary + '.');
    }
  } else {
    insights[2].items.push('Não foram observadas promoções no período selecionado.');
  }

  if(removed.length) {
    insights[3].items.push(removed.length + ' referência(s) observadas durante o mês não estavam presentes no último dia observado (' + dates[dates.length-1] + '). Tratar como sinal online a validar, não como confirmação de saída de loja física.');
  } else {
    insights[3].items.push('Não foram identificadas ausências relevantes no fim do período observado no site.');
  }
  insights[3].items.push('Confirmar impacto com dados de sell-out antes de retirar conclusões sobre performance de mercado.');

  return insights;
}

function renderMonthlyInsights(sections) {
  return '<div class="insight-grid">' + sections.map(sec =>
    '<div class="insight-card">' +
      '<h3>'+esc(sec.title)+'</h3>' +
      '<ul>' + sec.items.map(item => '<li>'+esc(item)+'</li>').join('') + '</ul>' +
    '</div>'
  ).join('') + '</div>';
}

function insightsToText(sections) {
  return sections.map(sec => sec.title + ':\n' + sec.items.map(item => '- ' + item).join('\n')).join('\n\n');
}

function shortProductName(name) {
  return (name || '').replace(/\s+/g,' ').trim().substring(0,70);
}

function copyMonthlyText() {
  const el = document.getElementById('monthly-copy');
  const text = el.value;
  try {
    navigator.clipboard.writeText(text);
    alert('Insights copiados.');
  } catch(e) {
    el.focus();
    el.select();
    document.execCommand('copy');
    alert('Insights copiados.');
  }
}

function downloadMonthlyExcel() {
  if(!monthlyData) return;
  const wb = XLSX.utils.book_new();

  const novidades = [['Categoria','Marca','Produto','Primeiro dia','Preço','Desconto %','Stock','URL']];
  monthlyData.newProducts.forEach(x => novidades.push([x.prod.c,x.prod.m,x.prod.n,x.first,x.h.p||'',x.h.d||'',x.h.s?'Sem Stock':'Disponível',x.prod.u]));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(novidades), 'Novidades');

  const promos = [['Categoria','Marca','Produto','Dias promo','Max desconto %','Preco minimo','Primeira promo observada','URL']];
  monthlyData.promos.forEach(x => promos.push([x.prod.c,x.prod.m,x.prod.n,x.days,x.maxDisc,x.minPrice||'',x.newPromo?'Sim':'Não',x.prod.u]));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(promos), 'Promocoes');

  const _covXls = (x,n) => x.promoRefs && n ? Math.round(x.promoDays/x.promoRefs/n*100)/100 : '';
  const _ndates = monthlyData.dates.length;
  const marcas = [['Marca','Refs ativas','Novidades','Refs em promo','Cobertura promo','Desconto medio %','Max desconto %']];
  monthlyData.brandRows.forEach(x => marcas.push([x.brand,x.active,x.newCount,x.promoRefs,_covXls(x,_ndates),x.discObs?Number((x.discSum/x.discObs).toFixed(1)):'',x.maxDisc||'']));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(marcas), 'Marcas');

  const categorias = [['Categoria','Refs ativas','Novidades','Refs em promo','Cobertura promo','Desconto medio %','Max desconto %']];
  monthlyData.categoryRows.forEach(x => categorias.push([x.category,x.active,x.newCount,x.promoRefs,_covXls(x,_ndates),x.discObs?Number((x.discSum/x.discObs).toFixed(1)):'',x.maxDisc||'']));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(categorias), 'Categorias');

  const ausentes = [['Categoria','Marca','Produto','Ultimo dia visto','Ultimo preco','Ultimo desconto %','URL']];
  monthlyData.removed.forEach(x => ausentes.push([x.prod.c,x.prod.m,x.prod.n,x.last.date,x.last.h.p||'',x.last.h.d||'',x.prod.u]));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ausentes), 'Presenca online');

  XLSX.writeFile(wb, 'Relatorio_Wells_' + monthlyData.month + '.xlsx');
}

// ============================================================
// UTILITARIOS
// ============================================================
function fmt(v) { return v!=null ? '€'+Number(v).toFixed(2) : '-'; }
function esc(s) { return (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function card(val, lbl) { return '<div class="card"><div class="val">'+val+'</div><div class="lbl">'+lbl+'</div></div>'; }

function badgeDestaque(t) {
  if(!t) return '';
  const parts = t.split(';').map(s=>s.trim()).filter(Boolean);
  return parts.map(p => {
    if(p.includes('Best')) return '<span class="badge badge-bs">'+p+'</span>';
    if(p.includes('Online')) return '<span class="badge badge-eo">'+p+'</span>';
    if(p.includes('Novo')) return '<span class="badge badge-novo">'+p+'</span>';
    return '<span class="badge" style="background:#eee;color:#555">'+p+'</span>';
  }).join(' ');
}

function renderPagination(containerId, pages, current, cb) {
  const div = document.getElementById(containerId);
  if(pages<=1) { div.innerHTML=''; return; }
  const cbStr = cb.toString();
  let btns = '';
  if(current>1) btns += '<button onclick="('+cbStr+')('+(current-1)+')">‹ Anterior</button>';
  const start = Math.max(1, current-2);
  const end   = Math.min(pages, current+2);
  if(start>1) btns += '<button onclick="('+cbStr+')(1)">1</button><span style="padding:4px 4px;color:#999">…</span>';
  for(let i=start;i<=end;i++) {
    btns += '<button class="'+(i===current?'active':'')+'" onclick="('+cbStr+')('+i+')">'+i+'</button>';
  }
  if(end<pages) btns += '<span style="padding:4px 4px;color:#999">…</span><button onclick="('+cbStr+')('+pages+')">'+pages+'</button>';
  if(current<pages) btns += '<button onclick="('+cbStr+')('+(current+1)+')">Próximo ›</button>';
  btns += '<span style="font-size:12px;color:#666;padding:4px 8px">Página </span>';
  btns += '<input type="number" min="1" max="'+pages+'" value="'+current+'" style="width:48px;padding:4px;border:1px solid #ccc;border-radius:4px;font-size:12px;text-align:center" ';
  btns += 'onchange="const n=Math.min('+pages+',Math.max(1,parseInt(this.value)||1));('+cbStr+')(n)">';
  btns += '<span style="font-size:12px;color:#666;padding:4px 4px">/ '+pages+'</span>';
  div.innerHTML = btns;
}

// Ordenacao de tabelas
function downloadExcel() {
  if(!evoSelected.length) return;
  const prods     = evoSelected.map(id => ALL_PRODUCTS.find(p => p.id===id)).filter(Boolean);
  const startDate = document.getElementById('evo-date-start').value;
  const endDate   = document.getElementById('evo-date-end').value;

  const rows = [['Produto','Marca','Categoria','Data','Preço','PVPR','Desconto %','Poupança €','Stock']];
  prods.forEach(prod => {
    Object.keys(prod.h).sort().filter(d => {
      if(startDate && d < startDate) return false;
      if(endDate   && d > endDate)   return false;
      return true;
    }).forEach(d => {
      const h = prod.h[d];
      rows.push([prod.n, prod.m, prod.c, d,
        h.p||'', h.v||'', h.d||0,
        (h.v&&h.p) ? parseFloat((h.v-h.p).toFixed(2)) : '',
        h.s ? 'Sem Stock' : 'Disponível']);
    });
  });

  if(rows.length <= 1) { alert('Nenhum registo no período selecionado.'); return; }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [40,20,20,12,10,10,12,12,12].map(w=>({wch:w}));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Histórico');
  const filename = prods.length===1
    ? (prods[0].m+' - '+prods[0].n).replace(/[\\/:*?"<>|]/g,'_')+'.xlsx'
    : 'Comparacao_Wells.xlsx';
  XLSX.writeFile(wb, filename);
}

</script>
</body>
</html>
'@

# Substituir os placeholders com os dados reais
$html = $html -replace 'DATES_JSON',    $jsonDates
$html = $html -replace 'CATS_JSON',     $jsonCats
$html = $html -replace 'BRANDS_JSON',   $jsonBrands
$html = $html -replace 'PRODUCTS_JSON', $jsonProducts
$html = $html -replace 'TIMESTAMP_VAL', $Timestamp

$html | Out-File -FilePath $DashboardOut -Encoding UTF8
Write-Host ("Dashboard gerado: " + $DashboardOut) -ForegroundColor Green
Write-Host ("  Produtos: " + $products.Count + " | Datas: " + $dates.Count) -ForegroundColor Green
