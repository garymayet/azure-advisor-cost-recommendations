# Azure Advisor Cost Recommendations — Multi-Subscription Report

Script en PowerShell que extrae las recomendaciones de **Azure Advisor** (categoría `Cost` por defecto) de todas las suscripciones especificadas dentro de un tenant, y genera un **CSV** consolidado + un **Excel con dashboard** (KPIs, gráficos de ahorro por suscripción y por impacto).

Cubre: resizing de VMs, shutdown de recursos ociosos, Reserved Instances, Savings Plans, cambios de SKU/familia, discos huérfanos, etc.

---

## 1. Requisitos

- **PowerShell 7+** (recomendado). Con Windows PowerShell 5.1 también funciona, pero 7 es más estable con `Az`.
- Cuenta con permisos de lectura (`Reader` o superior) sobre las suscripciones del tenant.
- Módulos:
  - `Az.Accounts` ≥ 2.12
  - `Az.Advisor` ≥ 2.0
  - `ImportExcel` ≥ **7.8** *(clave — versiones viejas no soportan `-ChartDefinition` y el dashboard falla)*

Verifica versiones:

```powershell
Get-Module Az.Accounts, Az.Advisor, ImportExcel -ListAvailable |
    Select-Object Name, Version | Sort-Object Name, Version -Descending
```

---

## 2. Instalación de módulos (una sola vez)

```powershell
Install-Module Az.Accounts, Az.Advisor, ImportExcel -Scope CurrentUser -Force -AllowClobber
```

Si ya los tenías instalados, actualízalos (evita el error `A parameter cannot be found that matches parameter name 'ChartDefinition'`):

```powershell
Update-Module Az.Accounts, Az.Advisor, ImportExcel -Force
```

Cierra y vuelve a abrir PowerShell después de instalar/actualizar.

---

## 3. Desbloquear el script

Si el archivo `.ps1` está en una carpeta sincronizada (Nextcloud, OneDrive, Dropbox) o lo descargaste de internet, Windows lo marca como bloqueado y verás:

> `File ... cannot be loaded. The file ... is not digitally signed.`

Solución:

```powershell
Unblock-File -Path .\Get-AzAdvisorCostRecommendations.ps1
```

Si además tu política de ejecución es `Restricted`, cámbiala a `RemoteSigned` para tu usuario (permite scripts locales):

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Alternativa sin cambiar la policy (solo esa sesión):

```powershell
powershell -ExecutionPolicy Bypass -File .\Get-AzAdvisorCostRecommendations.ps1 -TenantId "<TU_TENANT_GUID>"
```

---

## 4. Obtener el Tenant ID

Si no lo tienes a la mano:

```powershell
Connect-AzAccount
Get-AzTenant | Select-Object Id, Name, Domains
```

O sin login, si conoces un dominio del tenant:

```powershell
(Invoke-RestMethod "https://login.microsoftonline.com/<dominio>/v2.0/.well-known/openid-configuration").issuer
```

El GUID en el `issuer` es el tenant ID.

---

## 5. Ejecución

### Uso básico (guarda en `.\output` dentro de la carpeta del script)

```powershell
.\Get-AzAdvisorCostRecommendations.ps1 -TenantId "<TU_TENANT_GUID>"
```

### Guardar en la carpeta actual

```powershell
.\Get-AzAdvisorCostRecommendations.ps1 -TenantId "<TU_TENANT_GUID>" -OutputPath (Get-Location).Path
```

### Guardar en una ruta específica

```powershell
.\Get-AzAdvisorCostRecommendations.ps1 -TenantId "<TU_TENANT_GUID>" -OutputPath "C:\Reports\Advisor"
```

### Cambiar categoría

```powershell
.\Get-AzAdvisorCostRecommendations.ps1 -TenantId "<TU_TENANT_GUID>" -Category All
```

Valores válidos: `Cost` (default), `Performance`, `Security`, `HighAvailability`, `OperationalExcellence`, `All`.

### Filtrar a un subset de suscripciones

```powershell
.\Get-AzAdvisorCostRecommendations.ps1 -TenantId "<TU_TENANT_GUID>" `
    -SubscriptionIds "<SUB_GUID_1>","<SUB_GUID_2>"
```

Si no pasas `-SubscriptionIds`, usa la lista embebida en el script — edita el arreglo `$SubscriptionIds` y la tabla `$SubscriptionAlias` con tus propias suscripciones.

---

## 6. Salidas

En la carpeta indicada por `-OutputPath` se generan dos archivos con timestamp:

- `advisor_<Category>_YYYYMMDD_HHmm.csv` → datos crudos.
- `advisor_<Category>_YYYYMMDD_HHmm.xlsx` con 4 hojas:
  - **Dashboard** — KPIs (total de recomendaciones, ahorro anual USD, suscripciones evaluadas, fecha) + gráfico de barras de ahorro por suscripción y pie de impacto.
  - **Resumen_Suscripcion** — total, desglose Alto/Medio/Bajo y ahorro anual por sub.
  - **Resumen_Tipo** — agrupado por `RecommendationType` (útil para identificar patrones, p. ej. cuántas VMs candidatas a resize).
  - **Detalle** — una fila por recomendación con `ResourceId`, `TargetSku`, `Term`, `Region`, `AnnualSavingsUSD`, etc.

---

## 7. Uso en Power BI

1. Power BI Desktop → **Get Data → Text/CSV** → apunta al CSV generado.
2. Modela con `Subscription`, `Impact`, `RecommendationType`, `AnnualSavingsUSD`.
3. Sugerencias de visuales:
   - Tarjeta: suma total de `AnnualSavingsUSD`.
   - Barras: ahorro por `Subscription`.
   - Treemap: ahorro por `RecommendationType`.
   - Tabla: `ResourceId`, `TargetSku`, `Term`, `AnnualSavingsUSD` filtrable por sub.

Si quieres el `.pbit` (template con medidas DAX listas), avísame.

---

## 8. Troubleshooting

| Error | Causa | Solución |
|---|---|---|
| `not digitally signed` | Archivo bloqueado o execution policy restrictiva. | `Unblock-File` + `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`. |
| `Parameter set cannot be resolved using the specified named parameters` en `Get-AzAdvisorRecommendation` | Versión vieja de `Az.Advisor`. | `Update-Module Az.Advisor -Force` y reabrir PowerShell. El script ya trae fallback filtrando en memoria. |
| `A parameter cannot be found that matches parameter name 'ChartDefinition'` | `ImportExcel` viejo (< 7.x). | `Update-Module ImportExcel -Force`. |
| `Sin acceso a la suscripción <guid>` | Tu cuenta no tiene rol en esa sub. | Solicita `Reader` sobre la sub, o quítala del arreglo `-SubscriptionIds`. |
| `Connect-AzAccount` abre el navegador cada corrida | No hay contexto previo o el tenant es distinto. | Normal la primera vez; después reutiliza el contexto en la sesión. |
| Ahorros en 0 en muchas filas | Advisor no siempre expone `annualSavingsAmount` (típico en shutdown genérico). | Se conserva la recomendación; el ahorro real hay que estimarlo aparte. |

---

## 9. Notas operativas

- El script hace `Connect-AzAccount -Tenant` **solo si** no hay contexto o el tenant actual es distinto — evita re-login innecesario.
- Si una suscripción no tiene acceso, la salta con `WARNING` y sigue con la siguiente.
- Los importes de ahorro salen de `extendedProperties.annualSavingsAmount` cuando Advisor los expone (VM resize, shutdown, RI, savings plan). Cuando no vienen, queda en `0` pero la recomendación se conserva en el reporte.
- Recomendado correr con un usuario de servicio o con MFA cacheado si vas a agendarlo (Task Scheduler / Azure Automation).

---

## 10. Ejemplo de corrida limpia (paso a paso)

```powershell
# 1. Módulos
Install-Module Az.Accounts, Az.Advisor, ImportExcel -Scope CurrentUser -Force -AllowClobber

# 2. Desbloqueo + policy (una sola vez)
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
Unblock-File -Path .\Get-AzAdvisorCostRecommendations.ps1

# 3. Ejecución
.\Get-AzAdvisorCostRecommendations.ps1 `
    -TenantId "<TU_TENANT_GUID>" `
    -OutputPath (Get-Location).Path
```

Al terminar tendrás el CSV + XLSX en la carpeta desde donde ejecutaste.
