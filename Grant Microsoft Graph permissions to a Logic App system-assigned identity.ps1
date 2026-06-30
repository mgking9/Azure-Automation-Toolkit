# Adds certain MS Graph permissions to a system assigned managed identity in a specified Logic App.
#Requires -Modules "Az.Accounts", "Az.Resources", "Microsoft.Graph.Applications"

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param (
    [Parameter(Mandatory=$true)]
    [string]$LogicAppName,

    [Parameter(Mandatory=$true)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory=$true)]
    [string]$Tenant,

    [Parameter(Mandatory=$true)]
    [string]$Subscription
)

# ========================
# MICROSOFT GRAPH APP ID
# ========================
$GRAPH_APP_ID = "00000003-0000-0000-C000-000000000000"

# ========================
# CONNECT
# ========================
Connect-AzAccount -TenantId $Tenant -Subscription $Subscription | Out-Null
Connect-MgGraph -Identity -NoWelcome

Write-Host "AZ context"
Get-AzContext | Format-List

Write-Host "MG context"
Get-MgContext | Format-List

# ========================
# GRAPH PERMISSIONS YOU WANT
# ========================
$GraphPermissions = @(
    "EntitlementManagement.Read.All",
    "User.Read.All"
)

# ========================
# GET LOGIC APP MANAGED IDENTITY
# ========================
$logicAppResource = Get-AzResource `
    -ResourceGroupName $ResourceGroupName `
    -ResourceType "Microsoft.Logic/workflows" `
    -Name $LogicAppName `
    -ErrorAction Stop

$logicAppDetails = Get-AzResource `
    -ResourceId $logicAppResource.ResourceId `
    -ExpandProperties `
    -ErrorAction Stop

if (-not $logicAppDetails.Identity) {
    throw "Managed Identity is NOT enabled on Logic App."
}

$AutomationMSI = Get-AzADServicePrincipal -ObjectId $logicAppDetails.Identity.principalId

Write-Host "Target Logic App MSI:"
Write-Host "$($AutomationMSI.DisplayName) ($($AutomationMSI.Id))"

# ========================
# GET MICROSOFT GRAPH SP
# ========================
$GraphServicePrincipal = Get-AzADServicePrincipal -Filter "appId eq '$GRAPH_APP_ID'"

# ========================
# GET GRAPH APP ROLES
# ========================
$GraphAppRoles = $GraphServicePrincipal.AppRole | Where-Object {
    $_.Value -in $GraphPermissions -and $_.AllowedMemberType -contains "Application"
}

if ($GraphAppRoles.Count -ne $GraphPermissions.Count) {
    Write-Warning "Expected: $($GraphPermissions -join ', ')"
    Write-Warning "Found: $($GraphAppRoles.Value -join ', ')"
    throw "Some App Roles are missing on Microsoft Graph service principal"
}

# ========================
# ASSIGN ROLES (WhatIf-enabled)
# ========================
foreach ($AppRole in $GraphAppRoles) {

    if ($PSCmdlet.ShouldProcess(
        "Logic App MSI: $($AutomationMSI.DisplayName)",
        "Assign Graph permission: $($AppRole.Value)"
    )) {

        Write-Host "Assigning $($AppRole.Value)"

        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $AutomationMSI.Id `
            -PrincipalId $AutomationMSI.Id `
            -ResourceId $GraphServicePrincipal.Id `
            -AppRoleId $AppRole.Id | Out-Null
    }
}

Write-Host "DONE: Graph permission assignment completed."