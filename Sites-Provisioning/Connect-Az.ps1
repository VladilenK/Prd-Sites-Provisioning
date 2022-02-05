# before running the function locally

Connect-AzAccount

$adminUPN = "vkarass2@optumcloud.com"; $adminUPN | clip
$subsId = "df296ff9-4d03-4d99-ba49-098ca7209655"
#$subs1 = Get-AzSubscription -SubscriptionId $subsId
#Set-AzContext -SubscriptionObject $subs1
$TenantId = "db05faca-c82a-4b9d-b9c5-0f64b6755421"
Connect-AzAccount -subscription $subsId -Tenant $TenantId
Get-AzContext
Get-AzContext -ListAvailable
Get-AzResourceGroup | ft -a
#Clear-azcontext


#6V09tLW8PXVUuq-m.6n.dZH~_24.SNIPXG
