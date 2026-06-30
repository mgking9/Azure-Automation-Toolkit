#Removes certain MS Graph permissions to a system assigned managed identity in a specified Logic App.
#Requires -Modules "Az.Accounts", "Az.Resources", "Microsoft.Graph.Applications"

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
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
$GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"

# ========================
# GRAPH PERMISSIONS TO REMOVE
# ========================
$GraphPermissions = @(
    "EntitlementManagement.Read.All",
    "User.Read.All"
)

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

Write-Host "Removing permissions from Logic App: $($AutomationMSI.DisplayName) ($($AutomationMSI.Id))"

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

# ========================
# GET EXISTING ASSIGNMENTS
# ========================
$Assignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $AutomationMSI.Id |
    Where-Object {
        $_.ResourceId -eq $GraphServicePrincipal.Id
    }

if (-not $Assignments) {
    Write-Host "No Microsoft Graph permissions found on this Logic App."
    return
}

# ========================
# REMOVE MATCHING ROLES
# ========================
foreach ($AppRole in $GraphAppRoles) {

    $matchingAssignments = $Assignments | Where-Object {
        $_.AppRoleId -eq $AppRole.Id
    }

    foreach ($assignment in $matchingAssignments) {

        if ($PSCmdlet.ShouldProcess(
            "Logic App MSI: $($AutomationMSI.DisplayName)",
            "Remove Graph permission: $($AppRole.Value)"
        )) {

            Write-Host "Removing $($AppRole.Value)"

            # ========================
            # CORRECT REMOVE METHOD
            # ========================
            Remove-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $AutomationMSI.Id `
                -AppRoleAssignmentId $assignment.Id
        }
    }
}

Write-Host "DONE: Selected Graph permissions removed successfully."