# =============================================================================
# Migracao CSV -> SQLite  (executar uma unica vez)
# Importa wells_historico.csv para wells.db
# Uso: powershell -ExecutionPolicy Bypass -File scripts\migrar_para_sqlite.ps1
# =============================================================================

param(
    [string]$DataDir = "$(Split-Path $PSScriptRoot -Parent)\data"
)

$CsvPath = Join-Path $DataDir "historico\wells_historico.csv"
$DbPath  = Join-Path $DataDir "historico\wells.db"

# ---------------------------------------------------------------------------
# Instalar PSSQLite se necessario
# ---------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name PSSQLite)) {
    Write-Host "A instalar modulo PSSQLite..." -ForegroundColor Cyan
    Install-Module PSSQLite -Scope CurrentUser -Force -Repository PSGallery
}
Import-Module PSSQLite -ErrorAction Stop
Write-Host "PSSQLite carregado." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Verificar CSV
# ---------------------------------------------------------------------------
if (-not (Test-Path $CsvPath)) {
    Write-Host "CSV nao encontrado: $CsvPath" -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------------------
# Criar schema
# ---------------------------------------------------------------------------
Write-Host "A criar base de dados: $DbPath" -ForegroundColor Cyan

$schema = @"
CREATE TABLE IF NOT EXISTS historico (
    Data          TEXT    NOT NULL,
    Hora          TEXT,
    Categoria     TEXT    NOT NULL,
    ProdID        TEXT    NOT NULL,
    Marca         TEXT    NOT NULL,
    Produto       TEXT    NOT NULL,
    Preco         REAL,
    PVPR          REAL,
    Desconto_Pct  INTEGER,
    Poupanca_Euro REAL,
    Destaque      TEXT,
    Stock         TEXT,
    URL           TEXT,
    PRIMARY KEY (Data, ProdID)
);
CREATE INDEX IF NOT EXISTS idx_data      ON historico(Data);
CREATE INDEX IF NOT EXISTS idx_marca     ON historico(Marca);
CREATE INDEX IF NOT EXISTS idx_categoria ON historico(Categoria);
"@

Invoke-SqliteQuery -DataSource $DbPath -Query $schema

# ---------------------------------------------------------------------------
# Ler CSV
# ---------------------------------------------------------------------------
Write-Host "A ler CSV..." -ForegroundColor Cyan
$rows = Import-Csv -Path $CsvPath -Delimiter ";" -Encoding UTF8
Write-Host ("  " + $rows.Count + " linhas encontradas") -ForegroundColor Green

# ---------------------------------------------------------------------------
# Importar em transacao unica (performance)
# ---------------------------------------------------------------------------
$insertSql = @"
INSERT OR REPLACE INTO historico
    (Data, Hora, Categoria, ProdID, Marca, Produto,
     Preco, PVPR, Desconto_Pct, Poupanca_Euro, Destaque, Stock, URL)
VALUES
    (@Data, @Hora, @Categoria, @ProdID, @Marca, @Produto,
     @Preco, @PVPR, @Desconto_Pct, @Poupanca_Euro, @Destaque, @Stock, @URL)
"@

function Parse-Num   { param($s) if ($s -and $s -ne "") { try { [double]($s -replace ',','.') } catch { $null } } else { $null } }
function Parse-Int   { param($s) if ($s -and $s -ne "") { try { [int]$s } catch { $null } } else { $null } }

Write-Host "A importar dados (transacao unica)..." -ForegroundColor Cyan

$conn = New-SQLiteConnection -DataSource $DbPath
try {
    Invoke-SqliteQuery -SQLiteConnection $conn -Query "BEGIN TRANSACTION"
    $i = 0
    foreach ($row in $rows) {
        $params = @{
            Data          = $row.Data
            Hora          = $row.Hora
            Categoria     = $row.Categoria
            ProdID        = $row.ProdID
            Marca         = $row.Marca
            Produto       = $row.Produto
            Preco         = Parse-Num $row.Preco
            PVPR          = Parse-Num $row.PVPR
            Desconto_Pct  = Parse-Int $row.Desconto_Pct
            Poupanca_Euro = Parse-Num $row.Poupanca_Euro
            Destaque      = $row.Destaque
            Stock         = $row.Stock
            URL           = $row.URL
        }
        Invoke-SqliteQuery -SQLiteConnection $conn -Query $insertSql -SqlParameters $params
        $i++
        if ($i % 2000 -eq 0) { Write-Host ("  " + $i + " / " + $rows.Count + "...") -ForegroundColor Gray }
    }
    Invoke-SqliteQuery -SQLiteConnection $conn -Query "COMMIT"
    Write-Host "  Transacao confirmada." -ForegroundColor Green
} catch {
    Invoke-SqliteQuery -SQLiteConnection $conn -Query "ROLLBACK"
    Write-Host ("ERRO: " + $_) -ForegroundColor Red
    exit 1
} finally {
    $conn.Close()
}

# ---------------------------------------------------------------------------
# Verificar resultado
# ---------------------------------------------------------------------------
$total = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COUNT(*) AS n FROM historico").n
$datas = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT COUNT(DISTINCT Data) AS n FROM historico").n
$first = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT MIN(Data) AS d FROM historico").d
$last  = (Invoke-SqliteQuery -DataSource $DbPath -Query "SELECT MAX(Data) AS d FROM historico").d
$size  = [math]::Round((Get-Item $DbPath).Length / 1KB)

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Migracao concluida com sucesso!" -ForegroundColor Green
Write-Host ("  Registos : " + $total)          -ForegroundColor Green
Write-Host ("  Datas    : " + $datas + " (" + $first + " a " + $last + ")") -ForegroundColor Green
Write-Host ("  Tamanho  : " + $size + " KB")   -ForegroundColor Green
Write-Host ("  Ficheiro : " + $DbPath)          -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
