# Power BI - Azure Advisor Cost Dashboard

Como generar un .pbit funcional al 100% desde código es inviable (formato binario firmado por Microsoft), aquí tienes el paquete completo para armarlo en Power BI Desktop en menos de 5 minutos.

## Contenido del paquete

- `AdvisorCost_Query.pq` → M query lista para pegar en el Advanced Editor
- `AdvisorCost_Measures.dax` → 15 medidas DAX ya escritas
- `README_PowerBI.md` → estas instrucciones

## Paso a paso

### 1. Crear parámetro para la ruta del CSV

Power BI Desktop → **Home → Transform data → Manage Parameters → New**

- Name: `CsvPath`
- Type: Text
- Current Value: `C:\Reports\Advisor\advisor_Cost_latest.csv`
  (la ruta del CSV generado por el script; usa el `advisor_Cost_latest.csv` para apuntar siempre a la última corrida)

### 2. Crear la consulta

**Home → Get data → Blank Query** → **Advanced Editor** → pega el contenido de `AdvisorCost_Query.pq` → **Done**.

Renombra la consulta a `Advisor`.

**Close & Apply**.

### 3. Cargar las medidas

Por cada bloque separado por línea en blanco en `AdvisorCost_Measures.dax`:
- Home → New Measure
- Pega el bloque completo (incluyendo el nombre antes del `=`)
- Enter

### 4. Layout sugerido del dashboard

**Página 1 - Executive Summary**

| Zona | Visual | Campo/Medida |
|------|--------|--------------|
| Arriba izq | Card | `Ahorro Anual (Formato)` |
| Arriba centro | Card | `Total Recomendaciones` |
| Arriba der | Card | `Suscripciones Evaluadas` |
| Fila 2 izq | Card | `Ahorro Reserved Instances` |
| Fila 2 centro | Card | `Ahorro Savings Plan` |
| Fila 2 der | Card | `Ahorro Right-Size` |
| Centro | Bar chart | Eje: `Subscription`, Valores: `Ahorro Anual Total`, ordenado desc |
| Centro der | Donut | Leyenda: `RecommendationName`, Valores: `Ahorro Anual Total` |
| Abajo | Tabla | `Subscription`, `RecommendationName`, `Impact`, `Term`, `Ahorro Anual Total` |
| Slicer | Slicer | `Impact` |
| Slicer | Slicer | `Subscription` |
| Slicer | Slicer | `RecommendationName` |

**Página 2 - Detalle por Recurso**

- Tabla: `Subscription`, `ResourceGroup`, `ResourceName`, `TargetSku`, `Region`, `TermClean`, `AnnualSavingsUSD`
- Slicers laterales: `Subscription`, `RecommendationName`, `TargetSku`

**Página 3 - Reserved Instances / Savings Plans**

Filtra la página con `RecommendationName in {"VM Reserved Instance", "Savings Plan", "Azure Files Reserved Instance", "App Service Reserved Instance"}`.

- Matrix: filas = `Subscription`, columnas = `TermClean`, valores = `Ahorro Anual Total`
- Bar chart: `Subscription` vs `Ahorro Anual Total` por Term

### 5. Guardar como template

File → **Export → Power BI Template (.pbit)**

Te pedirá una descripción. Al distribuir el .pbit, quien lo abra solo necesita capturar la ruta del CSV en el parámetro `CsvPath` y refresca.

## Refrescar con nuevos datos

Cada vez que corras el script PowerShell se genera un CSV con timestamp nuevo. Dos opciones:

**A) Sobrescribir ruta fija:** en el script, usa siempre el mismo nombre de archivo (`advisor_Cost_latest.csv`) para que el dashboard apunte a él sin cambios.

**B) Cambiar parámetro:** al abrir el .pbit, actualiza `CsvPath` con la ruta del último CSV.

Para automatizar el B, en el script agrega al final:

```powershell
Copy-Item $csvFile (Join-Path $OutputPath "advisor_Cost_latest.csv") -Force
```

Y el `CsvPath` del template apunta siempre a `advisor_Cost_latest.csv`.
