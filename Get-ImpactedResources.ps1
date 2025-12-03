<#
.SYNOPSIS
  Lists resources likely to be impacted by an Azure subscription transfer and exports an Excel workbook
  with tabs: Summary, Details, RBAC, Networking, RecreationPlan.

.DESCRIPTION
  Uses Azure Resource Graph (ARG) for Summary/Details/Networking and Az RBAC APIs for role assignments.
  Avoids KQL 'let' to sidestep parser issues. Handles pagination and scoping:
    - Tenant (all accessible subscriptions)
    - Management Group (subs resolved via ARG)
    - Specific Subscriptions

.PARAMETER Scope
  Tenant (default) | ManagementGroup | Subscriptions

.PARAMETER ManagementGroupId
  When -Scope ManagementGroup.

.PARAMETER SubscriptionIds
  When -Scope Subscriptions.

.PARAMETER PageSize
  ARG page size (default 1000).

.PARAMETER OutputDirectory
  Where to write the Excel (and optional CSVs).

.PARAMETER AlsoExportCsv
  Additionally export per-tab CSVs.

.EXAMPLES
  .\Get-ImpactedResources.ps1
  .\Get-ImpactedResources.ps1 -Scope ManagementGroup -ManagementGroupId "contoso-mg"
  .\Get-ImpactedResources.ps1 -Scope Subscriptions -SubscriptionIds "sub-1","sub-2" -AlsoExportCsv
#>

[CmdletBinding()]
param(
    [ValidateSet('Tenant', 'ManagementGroup', 'Subscriptions')]
    [string]$Scope = 'Tenant',

    [string]$ManagementGroupId,

    [string[]]$SubscriptionIds,

    [int]$PageSize = 1000,

    [string]$OutputDirectory = (Get-Location).Path,

    [switch]$AlsoExportCsv
)

# --- Pre-flight checks ---
#Requires -Modules Az.Accounts, Az.ResourceGraph, Az.Resources

function Ensure-AzContext {
    if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
        Write-Error "Az.Accounts module is not installed. Run: Install-Module Az -Scope CurrentUser"
        exit 1
    }
    if (-not (Get-Module -ListAvailable -Name Az.ResourceGraph)) {
        Write-Error "Az.ResourceGraph module is not installed. Run: Install-Module Az.ResourceGraph -Scope CurrentUser"
        exit 1
    }

    try { $ctx = Get-AzContext -ErrorAction Stop } catch { $ctx = $null }
    if (-not $ctx) {
        Write-Host "Connecting to Azure..." -ForegroundColor Cyan
        Connect-AzAccount -ErrorAction Stop | Out-Null
    }
}

function Ensure-ImportExcel {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "Installing ImportExcel module (CurrentUser)..." -ForegroundColor Cyan
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module ImportExcel -ErrorAction Stop
}

function Invoke-ArgPaged {
    param(
        [Parameter(Mandatory)] [string] $Query,
        [Parameter(Mandatory)] [hashtable] $ScopeParams,
        [int] $PageSize = 1000
    )

    $all = @()
    $skip = $null

    do {
        $args = @{
            Query = $Query
            First = $PageSize
        } + $ScopeParams

        if ($skip) { $args['SkipToken'] = $skip }

        $resp = Search-AzGraph @args
        if ($resp -and $resp.Data) { $all += $resp.Data }

        $skip = $resp.SkipToken
    } while ($skip)

    return $all
}

function Get-ArgScopeParams {
    param(
        [string]$Scope,
        [string]$ManagementGroupId,
        [string[]]$SubscriptionIds
    )
    switch ($Scope) {
        'Tenant' {
            $subs = Get-AzSubscription | Where-Object { $_.State -eq 'Enabled' } | Select-Object -ExpandProperty Id
            if (-not $subs) { throw "No accessible subscriptions found." }
            return @{ Subscription = $subs }
        }
        'ManagementGroup' {
            if (-not $ManagementGroupId) { throw "ManagementGroupId is required for Scope=ManagementGroup." }
            return @{ ManagementGroup = $ManagementGroupId }
        }
        'Subscriptions' {
            if (-not $SubscriptionIds -or $SubscriptionIds.Count -eq 0) {
                throw "At least one SubscriptionId is required for Scope=Subscriptions."
            }
            return @{ Subscription = $SubscriptionIds }
        }
    }
}

function Resolve-SubscriptionIdsForRbac {
    param(
        [string]$Scope,
        [string]$ManagementGroupId,
        [string[]]$SubscriptionIds,
        [hashtable]$ScopeParams
    )

    switch ($Scope) {
        'Tenant' { return $ScopeParams.Subscription }
        'Subscriptions' { return $SubscriptionIds }
        'ManagementGroup' {
            $q = @"
ResourceContainers
| where type =~ 'microsoft.resources/subscriptions'
| project subscriptionId
"@
            $rows = Invoke-ArgPaged -Query $q -ScopeParams @{ ManagementGroup = $ManagementGroupId } -PageSize 1000
            $ids = $rows | Select-Object -ExpandProperty subscriptionId -Unique
            if (-not $ids) { throw "No subscriptions resolved under management group '$ManagementGroupId'." }
            return $ids
        }
    }
}

# Helpers
function Get-ResourceTypeFromId {
    param([string]$ResourceId)
    if ([string]::IsNullOrWhiteSpace($ResourceId)) { return $null }
    $idx = $ResourceId.IndexOf('/providers/', [System.StringComparison]::OrdinalIgnoreCase)
    if ($idx -lt 0) { return $null }
    $typePath = $ResourceId.Substring($idx + 11).Trim('/')
    $parts = $typePath.Split('/')
    if ($parts.Length -ge 2) { return ($parts[0] + '/' + $parts[1]).ToLower() }
    return $null
}
function Classify-ScopeLevel {
    param([string]$Scope)
    if ($Scope -match '/providers/[^/]+/[^/]+/') { return 'Resource' }
    elseif ($Scope -match '/resourceGroups/') { return 'ResourceGroup' }
    else { return 'Subscription' }
}
function Parse-SubscriptionIdFromScope {
    param([string]$Scope)
    if ($Scope -match '/subscriptions/([0-9a-fA-F-]+)') { return $matches[1] } else { return $null }
}
function Parse-ResourceGroupFromScope {
    param([string]$Scope)
    if ($Scope -match '/resourceGroups/([^/]+)') { return $matches[1] } else { return $null }
}

# --- KQL queries (no 'let') ---

# Impacted resource types (base + extra)
$impactedTypes = @(
    'microsoft.managedidentity/userassignedidentities',
    'microsoft.keyvault/vaults',
    'microsoft.sql/servers/databases',
    'microsoft.datalakestore/accounts',
    'microsoft.containerservice/managedclusters',
    'microsoft.storage/storageaccounts',
    'microsoft.web/sites',
    'microsoft.web/sites/slots',
    'microsoft.compute/virtualmachines',
    'microsoft.compute/disks',
    'microsoft.compute/snapshots',
    'microsoft.synapse/workspaces',
    'microsoft.purview/accounts',
    'microsoft.databricks/workspaces',
    'microsoft.eventhub/namespaces',
    'microsoft.servicebus/namespaces',
    'microsoft.network/privateendpoints',
    'microsoft.network/privatednszones',
    'microsoft.containerregistry/registries',
    'microsoft.logic/workflows',
    'microsoft.datafactory/factories',
    'microsoft.machinelearningservices/workspaces',
    'microsoft.app/managedenvironments',
    'microsoft.app/containerapps',
    'microsoft.cache/redis',
    'microsoft.kusto/clusters'
)
$impactedTypesLit = ($impactedTypes | ForEach-Object { "'$_'" }) -join ", "

$querySummary = @"
Resources
| where type in~ ($impactedTypesLit)
   or (isnotnull(identity) and tostring(identity.type) has 'SystemAssigned')
   or (type =~ 'microsoft.storage/storageaccounts' and tobool(properties.isHnsEnabled) == true)
| summarize count() by type
| order by count_ desc
"@

$queryDetails = @"
Resources
| where type in~ ($impactedTypesLit)
   or (isnotnull(identity) and tostring(identity.type) has 'SystemAssigned')
   or (type =~ 'microsoft.storage/storageaccounts' and tobool(properties.isHnsEnabled) == true)
| extend identityType = tostring(identity.type)
| extend isHnsEnabled = iif(type =~ 'microsoft.storage/storageaccounts', tobool(properties.isHnsEnabled), bool(null))
| extend diskEncryptionType = iif(type =~ 'microsoft.compute/disks', tostring(properties.encryption.type), '')
| extend diskKeyUrl        = iif(type =~ 'microsoft.compute/disks', tostring(properties.encryption.keyVaultProperties.keyUrl), '')
| extend vmHasEncryptionSettings = iif(type =~ 'microsoft.compute/virtualmachines', isnotempty(tostring(properties.storageProfile.osDisk.encryptionSettings)), bool(null))
| extend kvSoftDelete      = iif(type =~ 'microsoft.keyvault/vaults', tobool(properties.enableSoftDelete) or tobool(properties.enableSoftDeleteForPurgeProtection), bool(null))
| extend kvPurgeProtection = iif(type =~ 'microsoft.keyvault/vaults', tobool(properties.enablePurgeProtection), bool(null))
| project id, name, type, subscriptionId, resourceGroup, location,
          identityType, isHnsEnabled,
          diskEncryptionType, diskKeyUrl, vmHasEncryptionSettings,
          kvSoftDelete, kvPurgeProtection
| order by type asc, name asc
"@

# Networking queries
$queryPrivateEndpoints = @"
Resources
| where type =~ 'microsoft.network/privateendpoints'
| extend subnetId = tostring(properties.subnet.id)
| extend vnetId = iif(isnotempty(subnetId), tostring(split(subnetId, '/subnets/')[0]), '')
| extend subnetName = iif(isnotempty(subnetId), tostring(split(subnetId, '/subnets/')[1]), '')
| extend peId = id, peName = name
| mv-expand plc = properties.privateLinkServiceConnections to typeof(dynamic)
| extend targetId = tostring(plc.privateLinkServiceId)
| mv-expand gid = plc.groupIds
| summarize groupIds = make_set(tostring(gid), 50) by peId, peName, subscriptionId, resourceGroup, location, vnetId, subnetName, targetId
| extend groupIds = tostring(groupIds)
| project peId, peName, subscriptionId, resourceGroup, location, vnetId, subnetName, targetId, groupIds
"@

$queryPrivateDnsZoneGroups = @"
Resources
| where type =~ 'microsoft.network/privateendpoints/privateDnsZoneGroups'
| mv-expand zc = properties.privateDnsZoneConfigs to typeof(dynamic)
| extend peId = tostring(split(id, '/privateDnsZoneGroups/')[0])
| extend zoneId = tostring(zc.privateDnsZoneId)
| project peId, zoneId
"@

$queryPrivateDnsZones = @"
Resources
| where type =~ 'microsoft.network/privatednszones'
| project zoneId=id, zoneName=name
"@

$queryPrivateDnsVnetLinks = @"
Resources
| where type =~ 'microsoft.network/privatednszones/virtualnetworklinks'
| extend zoneId = tostring(split(id, '/virtualNetworkLinks/')[0])
| extend vnetId = tostring(properties.virtualNetwork.id)
| extend registrationEnabled = tobool(properties.registrationEnabled)
| project linkId=id, linkName=name, subscriptionId, resourceGroup, location, zoneId, vnetId, registrationEnabled
"@

# --- RUN ---
try {
    Ensure-AzContext
    $scopeParams = Get-ArgScopeParams -Scope $Scope -ManagementGroupId $ManagementGroupId -SubscriptionIds $SubscriptionIds

    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $excelPath = Join-Path $OutputDirectory "ImpactedResources_$stamp.xlsx"

    Write-Host "Querying Azure Resource Graph ($Scope scope)..." -ForegroundColor Cyan

    # 1) Summary & Details
    $summaryRows = Invoke-ArgPaged -Query $querySummary -ScopeParams $scopeParams -PageSize $PageSize
    $detailsRows = Invoke-ArgPaged -Query $queryDetails -ScopeParams $scopeParams -PageSize $PageSize

    # 2) Networking (split queries + PS joins)
    $peRows = Invoke-ArgPaged -Query $queryPrivateEndpoints       -ScopeParams $scopeParams -PageSize $PageSize
    $dzgRows = Invoke-ArgPaged -Query $queryPrivateDnsZoneGroups   -ScopeParams $scopeParams -PageSize $PageSize
    $zonesRows = Invoke-ArgPaged -Query $queryPrivateDnsZones        -ScopeParams $scopeParams -PageSize $PageSize
    $dnsLinkRows = Invoke-ArgPaged -Query $queryPrivateDnsVnetLinks    -ScopeParams $scopeParams -PageSize $PageSize

    # Zone name map & PE->zones
    $zoneNameById = @{}
    foreach ($z in $zonesRows) { if ($z.zoneId) { $zoneNameById[$z.zoneId] = $z.zoneName } }

    $zoneIdsByPe = @{}
    foreach ($g in $dzgRows) {
        if (-not $g.peId) { continue }
        if (-not $zoneIdsByPe.ContainsKey($g.peId)) { $zoneIdsByPe[$g.peId] = New-Object System.Collections.Generic.List[string] }
        if ($g.zoneId) { [void]$zoneIdsByPe[$g.peId].Add($g.zoneId) }
    }

    # Combine Networking
    $networkingCombined = @()

    foreach ($p in $peRows) {
        $targetType = Get-ResourceTypeFromId -ResourceId $p.targetId
        $zones = if ($zoneIdsByPe.ContainsKey($p.peId)) { $zoneIdsByPe[$p.peId] } else { @() }

        if ($zones.Count -gt 0) {
            foreach ($zid in $zones) {
                $networkingCombined += [PSCustomObject]@{
                    recordType     = 'PrivateEndpoint'
                    subscriptionId = $p.subscriptionId
                    resourceGroup  = $p.resourceGroup
                    location       = $p.location
                    peName         = $p.peName
                    peId           = $p.peId
                    vnetId         = $p.vnetId
                    subnetName     = $p.subnetName
                    targetId       = $p.targetId
                    targetType     = $targetType
                    dnsZoneId      = $zid
                    dnsZoneName    = $(if ($zoneNameById.ContainsKey($zid)) { $zoneNameById[$zid] } else { $null })
                    groupIds       = $(try { ( $p.groupIds | ConvertFrom-Json ) -join ';' } catch { $p.groupIds })
                }
            }
        }
        else {
            $networkingCombined += [PSCustomObject]@{
                recordType     = 'PrivateEndpoint'
                subscriptionId = $p.subscriptionId
                resourceGroup  = $p.resourceGroup
                location       = $p.location
                peName         = $p.peName
                peId           = $p.peId
                vnetId         = $p.vnetId
                subnetName     = $p.subnetName
                targetId       = $p.targetId
                targetType     = $targetType
                dnsZoneId      = $null
                dnsZoneName    = $null
                groupIds       = $(try { ( $p.groupIds | ConvertFrom-Json ) -join ';' } catch { $p.groupIds })
            }
        }
    }

    foreach ($d in $dnsLinkRows) {
        $networkingCombined += [PSCustomObject]@{
            recordType          = 'PrivateDnsVnetLink'
            subscriptionId      = $d.subscriptionId
            resourceGroup       = $d.resourceGroup
            location            = $d.location
            linkName            = $d.linkName
            linkId              = $d.linkId
            vnetId              = $d.vnetId
            dnsZoneId           = $d.zoneId
            dnsZoneName         = $(if ($zoneNameById.ContainsKey($d.zoneId)) { $zoneNameById[$d.zoneId] } else { $null })
            registrationEnabled = $d.registrationEnabled
            peName              = $null
            peId                = $null
            subnetName          = $null
            targetId            = $null
            targetType          = $null
            groupIds            = $null
        }
    }

    # 3) RBAC via Az.Resources
    Write-Host "Collecting RBAC role assignments (this can take several minutes in large tenants)..." -ForegroundColor Cyan
    $subsForRbac = Resolve-SubscriptionIdsForRbac -Scope $Scope -ManagementGroupId $ManagementGroupId -SubscriptionIds $SubscriptionIds -ScopeParams $scopeParams

    $rbacRaw = @()
    foreach ($sid in $subsForRbac) {
        Write-Host "  RBAC: $sid" -ForegroundColor DarkCyan
        try {
            $ra = Get-AzRoleAssignment -Scope "/subscriptions/$sid" -ErrorAction Stop
            if ($ra) { $rbacRaw += $ra }
        }
        catch {
            Write-Warning "Failed to get RBAC for subscription $sid. $_"
        }
    }

    # Normalize RBAC rows
    $rbacRows = foreach ($r in $rbacRaw) {
        $scopeLevel = Classify-ScopeLevel -Scope $r.Scope
        $subIdFromScope = Parse-SubscriptionIdFromScope -Scope $r.Scope
        $rgFromScope = Parse-ResourceGroupFromScope -Scope $r.Scope
        [PSCustomObject]@{
            id               = $r.Id
            scope            = $r.Scope
            scopeLevel       = $scopeLevel
            subscriptionId   = $subIdFromScope
            resourceGroup    = $rgFromScope
            principalId      = $r.ObjectId
            principalType    = $r.ObjectType
            principalName    = $r.DisplayName
            roleDefinitionId = $r.RoleDefinitionId
            roleName         = $r.RoleDefinitionName
            canDelegate      = $false
        }
    }

    # 4) Build RECREATION PLAN
    $recreationPlan = New-Object System.Collections.Generic.List[object]

    # RBAC recreation commands
    foreach ($r in $rbacRows) {
        $roleSpecifier = if ($r.roleName) { "-RoleDefinitionName `"$($r.roleName)`"" } else { "-RoleDefinitionId $($r.roleDefinitionId)" }
        $principalLabel = if ($r.principalName) { $r.principalName } else { $r.principalId }
        $cmd = "New-AzRoleAssignment -ObjectId $($r.principalId) $roleSpecifier -Scope `"$($r.scope)`""

        $recreationPlan.Add([PSCustomObject]@{
                category    = 'RBAC'
                scopeLevel  = $r.scopeLevel
                scope       = $r.scope
                target      = $principalLabel
                role        = $(if ($r.roleName) { $r.roleName } else { $r.roleDefinitionId })
                description = 'Recreate role assignment'
                command     = $cmd
            })
    }

    # Networking: Private DNS VNet links — PowerShell commands
    foreach ($n in $networkingCombined | Where-Object { $_.recordType -eq 'PrivateDnsVnetLink' }) {
        $zoneName = if ($n.dnsZoneName) { $n.dnsZoneName } elseif ($n.dnsZoneId -match '/privateDnsZones/([^/]+)') { $matches[1] } else { $null }
        $reg = if ($n.registrationEnabled) { '$true' } else { '$false' }
        $cmd = if ($zoneName) {
            "New-AzPrivateDnsVirtualNetworkLink -ResourceGroupName `"$($n.resourceGroup)`" -ZoneName `"$zoneName`" -Name `"$($n.linkName)`" -VirtualNetworkId `"$($n.vnetId)`" -RegistrationEnabled $reg"
        }
        else {
            "# Missing zone name for $($n.dnsZoneId) — resolve then run: New-AzPrivateDnsVirtualNetworkLink -ResourceGroupName `"$($n.resourceGroup)`" -ZoneName `"<zoneName>`" -Name `"$($n.linkName)`" -VirtualNetworkId `"$($n.vnetId)`" -RegistrationEnabled $reg"
        }

        $recreationPlan.Add([PSCustomObject]@{
                category    = 'Networking'
                scopeLevel  = 'Resource'
                scope       = $n.dnsZoneId
                target      = $n.vnetId
                role        = ''
                description = 'Ensure Private DNS zone is linked to VNet'
                command     = $cmd
            })
    }

    # Networking: PE ↔ Private DNS Zone association — action line (PE DNS Zone Group)
    foreach ($n in $networkingCombined | Where-Object { $_.recordType -eq 'PrivateEndpoint' -and $_.dnsZoneId }) {
        $action = "Associate PE `"$($n.peName)`" in RG `"$($n.resourceGroup)`" with Private DNS zone `"$($n.dnsZoneName)`" ($($n.dnsZoneId)) via a DNS Zone Group."
        $hint = "# Example (CLI): az network private-endpoint dns-zone-group create -g `"$($n.resourceGroup)`" --endpoint-name `"$($n.peName)`" -n default --private-dns-zone `"$($n.dnsZoneId)`""
        $recreationPlan.Add([PSCustomObject]@{
                category    = 'Networking'
                scopeLevel  = 'Resource'
                scope       = $n.peId
                target      = $n.dnsZoneId
                role        = ''
                description = $action
                command     = $hint
            })
    }

    # --- Write Excel workbook ---
    Ensure-ImportExcel
    $xlParams = @{
        Path         = $excelPath
        AutoSize     = $true
        FreezeTopRow = $true
        AutoFilter   = $true
        TableStyle   = 'Medium2'
    }

    if ($summaryRows) { $summaryRows | Export-Excel @xlParams -WorksheetName 'Summary'   -ClearSheet -TableName "SummaryTbl" }
    else { @()          | Export-Excel @xlParams -WorksheetName 'Summary'   -ClearSheet -TableName "SummaryTbl" }

    if ($detailsRows) { $detailsRows | Export-Excel @xlParams -WorksheetName 'Details'   -ClearSheet -TableName "DetailsTbl" }
    else { @()          | Export-Excel @xlParams -WorksheetName 'Details'   -ClearSheet -TableName "DetailsTbl" }

    if ($rbacRows) { $rbacRows    | Export-Excel @xlParams -WorksheetName 'RBAC'      -ClearSheet -TableName "RBACTbl" }
    else { @()          | Export-Excel @xlParams -WorksheetName 'RBAC'      -ClearSheet -TableName "RBACTbl" }

    if ($networkingCombined) { $networkingCombined | Export-Excel @xlParams -WorksheetName 'Networking' -ClearSheet -TableName "NetworkingTbl" }
    else { @()                 | Export-Excel @xlParams -WorksheetName 'Networking' -ClearSheet -TableName "NetworkingTbl" }

    if ($recreationPlan) { $recreationPlan | Export-Excel @xlParams -WorksheetName 'RecreationPlan' -ClearSheet -TableName "RecreationPlanTbl" }
    else { @()             | Export-Excel @xlParams -WorksheetName 'RecreationPlan' -ClearSheet -TableName "RecreationPlanTbl" }

    Write-Host "Excel workbook exported: $excelPath" -ForegroundColor Green

    if ($AlsoExportCsv) {
        $summaryCsv = Join-Path $OutputDirectory "ImpactedResources_Summary_$stamp.csv"
        $detailsCsv = Join-Path $OutputDirectory "ImpactedResources_Details_$stamp.csv"
        $rbacCsv = Join-Path $OutputDirectory "ImpactedResources_RBAC_$stamp.csv"
        $netCsv = Join-Path $OutputDirectory "ImpactedResources_Networking_$stamp.csv"
        $planCsv = Join-Path $OutputDirectory "ImpactedResources_RecreationPlan_$stamp.csv"

        if ($summaryRows) { $summaryRows        | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8 }
        if ($detailsRows) { $detailsRows        | Export-Csv -Path $detailsCsv -NoTypeInformation -Encoding UTF8 }
        if ($rbacRows) { $rbacRows           | Export-Csv -Path $rbacCsv    -NoTypeInformation -Encoding UTF8 }
        if ($networkingCombined) { $networkingCombined | Export-Csv -Path $netCsv     -NoTypeInformation -Encoding UTF8 }
        if ($recreationPlan) { $recreationPlan     | Export-Csv -Path $planCsv    -NoTypeInformation -Encoding UTF8 }

        Write-Host "CSV exports completed." -ForegroundColor Green
    }

    Write-Host "Done." -ForegroundColor Cyan
}
catch {
    Write-Error $_.Exception.Message
    throw
}
