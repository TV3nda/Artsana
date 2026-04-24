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

$CsvMaster = Join-Path $DataDir "historico\wells_historico.csv"
if (-not $DashboardOut) { $DashboardOut = Join-Path $DataDir "dashboard.html" }
$Timestamp    = Get-Date -Format "yyyy-MM-dd HH:mm"

if (-not (Test-Path $CsvMaster)) {
    Write-Host "Historico nao encontrado: $CsvMaster" -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------------------
# 1. Ler CSV e filtrar ultimos N dias
# ---------------------------------------------------------------------------
Write-Host "A ler historico..." -ForegroundColor Cyan
$rows = Import-Csv -Path $CsvMaster -Delimiter ";" -Encoding UTF8

$cutoff = (Get-Date).AddDays(-$DiasHistorico).ToString("yyyy-MM-dd")
$rows   = $rows | Where-Object { $_.Data -ge $cutoff }

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

    $preco   = if ($row.Preco)        { [decimal]($row.Preco -replace ',','.') }        else { $null }
    $pvpr    = if ($row.PVPR -and $row.PVPR -ne "") { [decimal]($row.PVPR -replace ',','.') } else { $null }
    $desc    = if ($row.Desconto_Pct -and $row.Desconto_Pct -ne "") { [int]$row.Desconto_Pct } else { $null }
    $poup    = if ($row.Poupanca_Euro -and $row.Poupanca_Euro -ne "") { [decimal]($row.Poupanca_Euro -replace ',','.') } else { $null }

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
<title>Wells Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;background:#f0f4f8;color:#333;font-size:14px}
header{background:#1a5276;color:#fff;padding:14px 24px;display:flex;align-items:center;justify-content:space-between}
header h1{font-size:20px;font-weight:bold}
header .meta{font-size:12px;opacity:.8}
nav{background:#2874a6;display:flex;gap:2px;padding:0 24px}
nav button{background:none;border:none;color:#cce;padding:12px 20px;cursor:pointer;font-size:14px;border-bottom:3px solid transparent}
nav button.active{color:#fff;border-bottom-color:#f39c12;font-weight:bold}
nav button:hover{color:#fff;background:rgba(255,255,255,.1)}
.tab{display:none;padding:20px 24px}
.tab.active{display:block}
.filters{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:16px;background:#fff;padding:14px;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.1)}
.filters label{font-size:12px;color:#555;display:flex;flex-direction:column;gap:3px}
.filters select,.filters input{padding:6px 10px;border:1px solid #ccc;border-radius:4px;font-size:13px;min-width:140px}
.filters input[type=text]{min-width:200px}
.filters input[type=checkbox]{width:16px;height:16px;margin-top:4px}
.btn{background:#2874a6;color:#fff;border:none;padding:8px 18px;border-radius:4px;cursor:pointer;font-size:13px}
.btn:hover{background:#1a5276}
.btn-compare{background:#e67e22}.btn-compare:hover{background:#ca6f1e}
table{width:100%;border-collapse:collapse;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.1);border-radius:8px;overflow:hidden}
th{background:#2874a6;color:#fff;padding:10px 8px;text-align:left;font-size:12px;white-space:nowrap;cursor:pointer;user-select:none}
th:hover{background:#1a5276}
th.sorted-asc::after{content:" ▲"}
th.sorted-desc::after{content:" ▼"}
td{padding:8px;border-bottom:1px solid #eee;font-size:12px}
tr:last-child td{border-bottom:none}
tr:hover td{background:#eaf4ff}
td a{color:#2980b9;text-decoration:none}
td a:hover{text-decoration:underline}
.badge{display:inline-block;padding:2px 7px;border-radius:10px;font-size:11px;font-weight:bold}
.badge-desc{background:#fdebd0;color:#c0392b}
.badge-bs{background:#fef9e7;color:#d4ac0d}
.badge-eo{background:#eaf2ff;color:#2471a3}
.badge-novo{background:#e9f7ef;color:#1e8449}
.stock-ok{color:#1e8449;font-weight:bold}
.stock-no{color:#c0392b;font-weight:bold}
.price-up{color:#c0392b;font-weight:bold}
.price-dn{color:#1e8449;font-weight:bold}
.price-eq{color:#888}
.summary-cards{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px}
.card{background:#fff;border-radius:8px;padding:14px 20px;box-shadow:0 1px 3px rgba(0,0,0,.1);min-width:140px;text-align:center}
.card .val{font-size:26px;font-weight:bold;color:#1a5276}
.card .lbl{font-size:11px;color:#888;margin-top:2px}
.chart-wrap{background:#fff;border-radius:8px;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,.1);margin-top:16px}
.chart-wrap canvas{max-height:320px}
.compare-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
.info{color:#666;font-style:italic;padding:20px;text-align:center}
#tbl-count{font-size:12px;color:#666;margin-bottom:8px}
.pagination{display:flex;gap:6px;margin-top:12px;align-items:center;flex-wrap:wrap}
.pagination button{padding:5px 10px;border:1px solid #ccc;background:#fff;border-radius:4px;cursor:pointer;font-size:12px}
.pagination button.active{background:#2874a6;color:#fff;border-color:#2874a6}
.pagination button:hover:not(.active){background:#eee}
.evo-cards-grid{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:8px}
.evo-card{display:flex;align-items:center;gap:10px;padding:9px 13px;background:#2c3e50;border:2px solid transparent;border-radius:8px;cursor:pointer;transition:all .18s;min-width:210px;max-width:300px;font-size:13px}
.evo-card:hover{background:#34495e}
.evo-card.selected{border-color:var(--ec,#2874a6);background:#1a252f}
.evo-dot{width:10px;height:10px;border-radius:50%;background:#555;flex-shrink:0;transition:background .18s}
.evo-card.selected .evo-dot{background:var(--ec,#2874a6)}
.evo-info{flex:1;min-width:0}
.evo-brand{font-size:11px;color:#aaa;display:block}
.evo-name{font-size:12px;color:#eee;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;display:block}
.evo-price{font-size:13px;font-weight:bold;color:#5dade2;white-space:nowrap;text-align:right}
.evo-sel-bar{display:flex;align-items:center;gap:12px;padding:9px 14px;background:#1a3a4a;border-radius:8px;margin-bottom:10px;font-size:13px;color:#ccc}
.evo-sel-bar span{flex:1}
.brand-grid{display:flex;gap:14px;align-items:flex-start;padding-bottom:12px}
#marcas-grid{overflow-x:auto}
.brand-scroll-bar{display:flex;align-items:center;gap:6px;margin-bottom:6px}
.brand-scroll-bar button{flex-shrink:0;background:#2874a6;color:#fff;border:none;border-radius:6px;width:32px;height:32px;font-size:20px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:background .15s}
.brand-scroll-bar button:hover{background:#1a5276}
#marcas-topscroll{flex:1;overflow-x:auto;overflow-y:hidden;height:14px}
.brand-col{min-width:190px;max-width:230px;flex-shrink:0;background:#fff;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.1);overflow:hidden}
.brand-hdr{background:#2874a6;color:#fff;padding:10px 14px;font-weight:bold;font-size:13px;display:flex;justify-content:space-between;align-items:center}
.brand-hdr .brand-count{font-size:11px;opacity:.8;font-weight:normal}
.brand-prod{padding:10px 14px;border-bottom:1px solid #f0f0f0;font-size:12px}
.brand-prod:last-child{border-bottom:none}
.brand-prod:hover{background:#f7fbff}
.brand-prod-name{display:block;color:#1a5276;font-weight:600;margin-bottom:5px;line-height:1.3;text-decoration:none}
.brand-prod-name:hover{text-decoration:underline}
.brand-prod-row{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.brand-price{color:#2874a6;font-weight:bold}
.brand-price-old{color:#999;text-decoration:line-through;font-size:11px}
</style>
</head>
<body>

<header>
  <div>
    <h1>Wells.pt — Dashboard de Preços</h1>
    <div class="meta">Categorias: Chupetas · Biberões · Bombas Tira Leite</div>
  </div>
  <div class="meta" id="hdr-meta"></div>
</header>

<nav>
  <button class="active" onclick="showTab('produtos')">Produtos</button>
  <button onclick="showTab('comparar')">Comparar Datas</button>
  <button onclick="showTab('evolucao')">Evolução de Preços</button>
  <button onclick="showTab('marcas')">Por Marca</button>
</nav>

<!-- ===== TAB: PRODUTOS ===== -->
<div id="tab-produtos" class="tab active">
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
  </div>
  <div class="summary-cards" id="cards-compare"></div>
  <div id="tbl-count-cmp"></div>
  <table id="tbl-compare">
    <thead>
      <tr>
        <th onclick="sortTable('tbl-compare',0)">Categoria</th>
        <th onclick="sortTable('tbl-compare',1)">Marca</th>
        <th onclick="sortTable('tbl-compare',2)">Produto</th>
        <th onclick="sortTable('tbl-compare',3)">Preço A</th>
        <th onclick="sortTable('tbl-compare',4)">Preço B</th>
        <th onclick="sortTable('tbl-compare',5)">Diferença €</th>
        <th onclick="sortTable('tbl-compare',6)">Variação %</th>
        <th onclick="sortTable('tbl-compare',7)">Desconto B</th>
        <th onclick="sortTable('tbl-compare',8)">Stock B</th>
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
  <div class="brand-scroll-bar">
    <button onclick="scrollMarcas(-1)" title="Anterior">&#8249;</button>
    <div id="marcas-topscroll"><div id="marcas-topscroll-inner" style="height:1px"></div></div>
    <button onclick="scrollMarcas(1)" title="Próximo">&#8250;</button>
  </div>
  <div id="marcas-grid"></div>
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
let currentPage = { produtos: 1, compare: 1 };
let evoChart = null;

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
  _syncTopScroll();
}

function scrollMarcas(dir) {
  const grid = document.getElementById('marcas-grid');
  if (!grid) return;
  grid.scrollBy({ left: dir * Math.round(grid.clientWidth * 0.75), behavior: 'smooth' });
}

function _syncTopScroll() {
  const inner = document.getElementById('marcas-topscroll-inner');
  const grid  = document.getElementById('marcas-grid');
  if (!inner || !grid) return;
  const bg = grid.querySelector('.brand-grid');
  inner.style.width = (bg ? bg.scrollWidth : grid.scrollWidth) + 'px';
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

  ['f-cat','cmp-cat','evo-cat','marca-cat'].forEach(id => {
    const sel = document.getElementById(id);
    CATS.forEach(c => { const o=document.createElement('option'); o.value=c; o.textContent=c; sel.appendChild(o); });
  });

  const fBrand = document.getElementById('f-brand');
  BRANDS.forEach(b => { const o=document.createElement('option'); o.value=b; o.textContent=b; fBrand.appendChild(o); });

  renderProducts();
  searchProducts();

  // Sincronização da barra de scroll dupla (topo <-> fundo) no tab Por Marca
  const _grid = document.getElementById('marcas-grid');
  const _top  = document.getElementById('marcas-topscroll');
  if (_grid && _top) {
    _grid.addEventListener('scroll', () => { _top.scrollLeft = _grid.scrollLeft; }, { passive:true });
    _top.addEventListener('scroll',  () => { _grid.scrollLeft = _top.scrollLeft;  }, { passive:true });
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
  const pages = Math.ceil(filtered.length / PAGE_SIZE);
  const p = currentPage.produtos;
  const slice = filtered.slice((p-1)*PAGE_SIZE, p*PAGE_SIZE);

  document.getElementById('tbl-count').textContent =
    filtered.length + ' produtos encontrados' + (pages>1 ? ' (página '+p+' de '+pages+')' : '');

  const tbody = document.getElementById('tbody-produtos');
  tbody.innerHTML = slice.map(prod => {
    const h = prod.h[date];
    const dCell = h.d ? '<span class="badge badge-desc">-'+h.d+'%</span>' : '-';
    const sCell = h.s ? '<span class="stock-no">Sem Stock</span>' : '<span class="stock-ok">Disponível</span>';
    return '<tr>' +
      '<td>'+esc(prod.c)+'</td>' +
      '<td>'+esc(prod.m)+'</td>' +
      '<td><a href="'+esc(prod.u)+'" target="_blank">'+esc(prod.n)+'</a></td>' +
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

  // Sort by abs diff desc by default
  rows.sort((a,b) => Math.abs(b.diff||0) - Math.abs(a.diff||0));

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
    const pctStr  = pct!==null  ? '<span class="'+cls+'">'+(pct>=0?'+':'')+pct.toFixed(1)+'%</span>'  : (type==='new'?'<span class="price-dn">Novo</span>':'<span class="price-up">Removido</span>');
    const dB  = hB?.d ? '<span class="badge badge-desc">-'+hB.d+'%</span>' : '-';
    const stB = hB ? (hB.s?'<span class="stock-no">Sem Stock</span>':'<span class="stock-ok">OK</span>') : '-';
    return '<tr>' +
      '<td>'+esc(prod.c)+'</td>' +
      '<td>'+esc(prod.m)+'</td>' +
      '<td><a href="'+esc(prod.u)+'" target="_blank">'+esc(prod.n)+'</a></td>' +
      '<td>'+prA+'</td><td>'+prB+'</td>' +
      '<td>'+diffStr+'</td><td>'+pctStr+'</td>' +
      '<td>'+dB+'</td><td>'+stB+'</td>' +
      '</tr>';
  }).join('');

  renderPagination('pag-compare', pages, p, (n)=>{ currentPage.compare=n; _renderComparePage(); });
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

function sortTable(tableId, colIdx) {
  const key = tableId + '_' + colIdx;
  const asc = sortState[key] !== 'asc';
  sortState[key] = asc ? 'asc' : 'desc';

  const table = document.getElementById(tableId);
  const ths = table.querySelectorAll('th');
  ths.forEach(th => { th.classList.remove('sorted-asc','sorted-desc'); });
  ths[colIdx].classList.add(asc ? 'sorted-asc' : 'sorted-desc');

  const tbody = table.querySelector('tbody');
  const rows  = Array.from(tbody.querySelectorAll('tr'));

  rows.sort((a, b) => {
    const aText = a.cells[colIdx]?.textContent?.trim() || '';
    const bText = b.cells[colIdx]?.textContent?.trim() || '';
    const aNum  = parseFloat(aText.replace(/[^0-9.\-]/g,''));
    const bNum  = parseFloat(bText.replace(/[^0-9.\-]/g,''));
    if(!isNaN(aNum) && !isNaN(bNum)) return asc ? aNum-bNum : bNum-aNum;
    return asc ? aText.localeCompare(bText,'pt') : bText.localeCompare(aText,'pt');
  });

  rows.forEach(r => tbody.appendChild(r));
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

# ---------------------------------------------------------------------------
# Publicar no GitHub Pages
# ---------------------------------------------------------------------------
$RepoDir = "C:\Claude Code\github\Artsana"
if (Test-Path $RepoDir) {
    try {
        # 1. Copiar o dashboard para a pasta do repositorio
        Copy-Item $DashboardOut (Join-Path $RepoDir "index.html") -Force

        # 2. Commit e push
        Set-Location $RepoDir
        git add index.html 2>&1 | Out-Null
        $commitMsg = "Dashboard atualizado - " + (Get-Date -Format "yyyy-MM-dd HH:mm")
        git commit -m $commitMsg 2>&1 | Out-Null
        git push origin master 2>&1 | Out-Null

        Write-Host "  Publicado em: https://tv3nda.github.io/Artsana" -ForegroundColor Cyan
    } catch {
        Write-Host ("  AVISO: Publicacao GitHub falhou - " + $_.Exception.Message) -ForegroundColor Red
    } finally {
        Set-Location $PSScriptRoot
    }
}
