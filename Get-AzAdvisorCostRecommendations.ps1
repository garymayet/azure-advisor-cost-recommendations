<#
.SYNOPSIS
    Extrae recomendaciones de Azure Advisor (categoría Cost) para múltiples
    suscripciones dentro de un tenant y genera un Excel consolidado con dashboard.

.REQUISITOS
    - PowerShell 7+
    - Módulos: Az.Accounts, Az.Advisor, ImportExcel
      Install-Module Az.Accounts, Az.Advisor, ImportExcel -Scope CurrentUser -Force

.USO
    .\Get-AzAdvisorCostRecommendations.ps1 -TenantId "<TU_TENANT_GUID>" -OutputPath "C:\Reports"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$TenantId,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter(Mandatory=$false)]
    [string[]]$SubscriptionIds = @(
        # Reemplaza con los GUIDs de tus suscripciones (o pásalos con -SubscriptionIds)
        "00000000-0000-0000-0000-000000000001",  # Suscripcion-A
        "00000000-0000-0000-0000-000000000002",  # Suscripcion-B
        "00000000-0000-0000-0000-000000000003"   # Suscripcion-C
    ),

    [ValidateSet('Cost','Performance','Security','HighAvailability','OperationalExcellence','All')]
    [string]$Category = 'Cost'
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }

# --- Etiquetas amigables (opcional: GUID -> nombre corto para los reportes) ---
$SubscriptionAlias = @{
    "00000000-0000-0000-0000-000000000001" = "Suscripcion-A"
    "00000000-0000-0000-0000-000000000002" = "Suscripcion-B"
    "00000000-0000-0000-0000-000000000003" = "Suscripcion-C"
}

# --- Login ---
$ctx = Get-AzContext -ErrorAction SilentlyContinue
if (-not $ctx -or $ctx.Tenant.Id -ne $TenantId) {
    Write-Host "Autenticando en tenant $TenantId ..." -ForegroundColor Cyan
    Connect-AzAccount -Tenant $TenantId | Out-Null
}

$all = New-Object System.Collections.Generic.List[object]

foreach ($subId in $SubscriptionIds) {
    $alias = if ($SubscriptionAlias.ContainsKey($subId)) { $SubscriptionAlias[$subId] } else { $subId }
    Write-Host "-> [$alias] $subId" -ForegroundColor Yellow
    try {
        Set-AzContext -SubscriptionId $subId -Tenant $TenantId | Out-Null
    } catch {
        Write-Warning "  Sin acceso a la suscripción $subId : $($_.Exception.Message)"
        continue
    }

    try {
        $recs = Get-AzAdvisorRecommendation
        if ($Category -ne 'All') {
            $recs = $recs | Where-Object { $_.Category -eq $Category }
        }
    } catch {
        Write-Warning "  Error obteniendo recomendaciones: $($_.Exception.Message)"
        continue
    }

    Write-Host "  $($recs.Count) recomendaciones" -ForegroundColor Green

    foreach ($r in $recs) {
        $ep       = $r.ExtendedProperty
        $savings  = 0.0
        $currency = ""
        $term     = ""
        $sku      = ""
        $region   = ""
        if ($ep) {
            if ($ep.ContainsKey('annualSavingsAmount'))   { [double]::TryParse($ep['annualSavingsAmount'],   [ref]$savings) | Out-Null }
            if ($ep.ContainsKey('savingsAmount'))         { if ($savings -eq 0) { [double]::TryParse($ep['savingsAmount'], [ref]$savings) | Out-Null } }
            if ($ep.ContainsKey('savingsCurrency'))       { $currency = $ep['savingsCurrency'] }
            if ($ep.ContainsKey('annualSavingsCurrency')) { if (-not $currency) { $currency = $ep['annualSavingsCurrency'] } }
            if ($ep.ContainsKey('term'))                  { $term = $ep['term'] }
            if ($ep.ContainsKey('targetSku'))             { $sku  = $ep['targetSku'] }
            if ($ep.ContainsKey('region'))                { $region = $ep['region'] }
        }

        $resourceId = $r.ResourceMetadataResourceId
        $rg = ""
        if ($resourceId -match "/resourceGroups/([^/]+)/") { $rg = $Matches[1] }

        $all.Add([PSCustomObject]@{
            Subscription        = $alias
            SubscriptionId      = $subId
            Category            = $r.Category
            Impact              = $r.Impact
            Problem             = $r.ShortDescriptionProblem
            Solution            = $r.ShortDescriptionSolution
            ResourceType        = ($resourceId -replace '.*/providers/([^/]+/[^/]+)/.*','$1')
            ResourceGroup       = $rg
            ResourceId          = $resourceId
            RecommendationType  = $r.RecommendationTypeId
            AnnualSavingsUSD    = [math]::Round($savings, 2)
            Currency            = $currency
            Term                = $term
            TargetSku           = $sku
            Region              = $region
            LastUpdated         = $r.LastUpdated
        })
    }
}

if ($all.Count -eq 0) {
    Write-Warning "No se obtuvieron recomendaciones."
    return
}

# --- Export CSV crudo ---
$stamp     = Get-Date -Format "yyyyMMdd_HHmm"
$csvFile   = Join-Path $OutputPath "advisor_${Category}_$stamp.csv"
$xlsxFile  = Join-Path $OutputPath "advisor_${Category}_$stamp.xlsx"
$csvLatest = Join-Path $OutputPath "advisor_${Category}_latest.csv"

$all | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
Write-Host "`nCSV: $csvFile" -ForegroundColor Cyan

# Copia con nombre fijo para que Power BI apunte siempre a la última corrida
Copy-Item $csvFile $csvLatest -Force
Write-Host "CSV (latest): $csvLatest" -ForegroundColor Cyan

# --- Excel con dashboard ---
Import-Module ImportExcel

# Hoja de detalle
$all | Export-Excel -Path $xlsxFile -WorksheetName "Detalle" `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName "tblDetalle"

# Resumen por suscripción
$bySub = $all | Group-Object Subscription | ForEach-Object {
    [PSCustomObject]@{
        Subscription   = $_.Name
        Total          = $_.Count
        Alto           = ($_.Group | Where-Object Impact -eq 'High').Count
        Medio          = ($_.Group | Where-Object Impact -eq 'Medium').Count
        Bajo           = ($_.Group | Where-Object Impact -eq 'Low').Count
        AhorroAnualUSD = [math]::Round((($_.Group | Measure-Object AnnualSavingsUSD -Sum).Sum), 2)
    }
} | Sort-Object AhorroAnualUSD -Descending

$bySub | Export-Excel -Path $xlsxFile -WorksheetName "Resumen_Suscripcion" `
    -AutoSize -AutoFilter -BoldTopRow -TableName "tblResumenSub"

# Resumen por tipo de recomendación
$byType = $all | Group-Object RecommendationType | ForEach-Object {
    [PSCustomObject]@{
        RecommendationType = $_.Name
        Ejemplo            = ($_.Group | Select-Object -First 1).Problem
        Ocurrencias        = $_.Count
        AhorroAnualUSD     = [math]::Round((($_.Group | Measure-Object AnnualSavingsUSD -Sum).Sum), 2)
    }
} | Sort-Object AhorroAnualUSD -Descending

$byType | Export-Excel -Path $xlsxFile -WorksheetName "Resumen_Tipo" `
    -AutoSize -AutoFilter -BoldTopRow -TableName "tblResumenTipo"

# --- Dashboard (compatible con versiones viejas de ImportExcel) ---
# KPIs en la hoja Dashboard
$totalRec   = $all.Count
$totalSave  = [math]::Round((($all | Measure-Object AnnualSavingsUSD -Sum).Sum), 2)
$subsEval   = ($all | Select-Object -Unique SubscriptionId).Count
$generado   = (Get-Date).ToString("yyyy-MM-dd HH:mm")

$kpis = @(
    [PSCustomObject]@{ Metrica = "Total recomendaciones";       Valor = $totalRec }
    [PSCustomObject]@{ Metrica = "Ahorro anual estimado (USD)"; Valor = $totalSave }
    [PSCustomObject]@{ Metrica = "Suscripciones evaluadas";     Valor = $subsEval }
    [PSCustomObject]@{ Metrica = "Generado";                    Valor = $generado }
)
$kpis | Export-Excel -Path $xlsxFile -WorksheetName "Dashboard" -StartRow 1 -AutoSize -BoldTopRow

# Impacto (tabla auxiliar en Dashboard, filas 8+)
$byImpact = $all | Group-Object Impact | ForEach-Object {
    [PSCustomObject]@{ Impacto = $_.Name; Cantidad = $_.Count }
} | Sort-Object Cantidad -Descending

# Chart 1: Ahorro anual por suscripción (usa la hoja Resumen_Suscripcion)
$chart1 = New-ExcelChartDefinition `
    -XRange "Resumen_Suscripcion!A2:A$($bySub.Count+1)" `
    -YRange "Resumen_Suscripcion!F2:F$($bySub.Count+1)" `
    -Title "Ahorro anual por suscripción (USD)" `
    -ChartType BarClustered -Row 8 -Column 3 -Width 700 -Height 400 -NoLegend

# Chart 2: Recomendaciones por impacto (pie sobre la tabla que insertamos abajo)
$firstDataRow = 9   # Dashboard!A9 empieza la data del pie
$lastDataRow  = 8 + $byImpact.Count
$chart2 = New-ExcelChartDefinition `
    -XRange "Dashboard!A$firstDataRow`:A$lastDataRow" `
    -YRange "Dashboard!B$firstDataRow`:B$lastDataRow" `
    -Title "Recomendaciones por impacto" `
    -ChartType Pie -Row 8 -Column 14 -Width 500 -Height 400

# Inserta el bloque de impacto + adjunta ambos charts al mismo Export-Excel
# (funciona con ImportExcel viejo, no requiere Add-ExcelChart)
$byImpact | Export-Excel -Path $xlsxFile -WorksheetName "Dashboard" `
    -StartRow 8 -AutoSize -BoldTopRow `
    -ExcelChartDefinition $chart1, $chart2

Write-Host "XLSX: $xlsxFile" -ForegroundColor Cyan
Write-Host "`nListo." -ForegroundColor Green
