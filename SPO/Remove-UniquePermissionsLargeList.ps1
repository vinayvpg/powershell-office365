#To call a non-generic method Load
Function Invoke-LoadMethod() {
    param(
            [Microsoft.SharePoint.Client.ClientObject]$Object = $(throw "Please provide a Client Object"),
            [string]$PropertyName
        ) 
   $ctx = $Object.Context
   $load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load") 
   $type = $Object.GetType()
   $clientLoad = $load.MakeGenericMethod($type)
  
   $Parameter = [System.Linq.Expressions.Expression]::Parameter(($type), $type.Name)
   $Expression = [System.Linq.Expressions.Expression]::Lambda([System.Linq.Expressions.Expression]::Convert([System.Linq.Expressions.Expression]::PropertyOrField($Parameter,$PropertyName),[System.Object] ), $($Parameter))
   $ExpressionArray = [System.Array]::CreateInstance($Expression.GetType(), 1)
   $ExpressionArray.SetValue($Expression, 0)
   $clientLoad.Invoke($ctx,@($Object,$ExpressionArray))
}

function Restore-SPOListAllItemsInheritance
{
   param (
        [Parameter(Mandatory=$true,Position=1)]
		[string]$Username,
		[Parameter(Mandatory=$true,Position=2)]
		[string]$Url,
        [Parameter(Mandatory=$true,Position=3)]
		[SecureString]$AdminPassword,
        [Parameter(Mandatory=$true,Position=4)]
		[string]$ListTitle
	)
  
    $ctx=New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username, $AdminPassword)
    $ctx.Load($ctx.Web.Lists)
    $ctx.Load($ctx.Web)
    $ctx.Load($ctx.Web.Webs)
    $ctx.ExecuteQuery()
    $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
    $ctx.Load($ll)
    $ctx.ExecuteQuery()

    ## View XML
    $qCommand = @"
        <View Scope='RecursiveAll'>
            <Query>
                <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
            </Query>
            <RowLimit Paged='TRUE'>5000</RowLimit>
        </View>
"@

    ## Page Position
    $position = $null
 
    ## All Items
    $allItems = @()
    do{
        $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
        $camlQuery.ListItemCollectionPosition = $position
        $camlQuery.ViewXml = $qCommand
        
        ## Executing the query
        $currentCollection = $ll.GetItems($camlQuery)
        $ctx.Load($currentCollection)
        $ctx.ExecuteQuery()
 
        ## Getting the position of the previous page
        $position = $currentCollection.ListItemCollectionPosition
 
        # Adding current collection to the allItems collection
        $allItems += $currentCollection

         Write-Host "Collecting items. Current number of items: " $allItems.Count
    }
    while($position -ne $null)
 
    Write-Host "Total number of items: " $allItems.Count

    $uniquePermissionsItemCount = 0
    $inheritPermissionsItemCount = 0;

    for($j=0; $j -lt $allItems.Count; $j++)
    {
        Invoke-LoadMethod -Object $allItems[$j] -PropertyName "HasUniqueRoleAssignments"
        Invoke-LoadMethod -Object $allItems[$j] -PropertyName "ContentType"
        $ctx.ExecuteQuery()

        $fileName = $allItems[$j]["Title"]
        $path = $allItems[$j]["FileRef"]
        $created = $allItems[$j]["Created"]
        $modified = $allItems[$j]["Modified"]
        $ctype = $allItems[$j].ContentType.Name

        "$fileName `t $path `t $created `t $modified `t $ctype" | out-file $global:logCSVPath -Append 

        if($allItems[$j].HasUniqueRoleAssignments -eq $true) {
            $uniquePermissionsItemCount += 1

            Write-Host "Resetting permissions for " $fileName "..." $path -NoNewline

            $allItems[$j].ResetRoleInheritance()
            $ctx.ExecuteQuery()

            Write-Host "...Done" -BackgroundColor Green
        }
        else {
            $inheritPermissionsItemCount += 1

            Write-Host "No unique permissions for " $fileName "..." $path -NoNewline
            Write-Host "...Skip" -BackgroundColor Magenta
        }
    }

    Write-Host "Total number of unique permissioned items: " $uniquePermissionsItemCount
    Write-Host "Total number of inherit permission items: " $inheritPermissionsItemCount
}

#------------------ main script --------------------------------------------------
cls

write-host "Start - $(Get-Date)`n" -ForegroundColor Yellow

$timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }
$currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$downloadPath = $currentDir + "\downloads\" + $timestamp + "\"

Write-Host "The log will be available at '$downloadPath'" -ForegroundColor Cyan

if(-not (Test-Path $downloadPath -PathType Container)) {
    Write-Host "Creating download folder '$downloadPath'..." -NoNewline
    md -Path $downloadPath | out-null
    Write-Host "Done" -ForegroundColor White -BackgroundColor Green
}

$global:logCSVPath = $global:downloadPath + "FullLog.csv"

#log csv
"FileName `t Path `t Created `t Modified `t CType" | out-file $global:logCSVPath

$siteUrl="https://murphyoil.sharepoint.com/sites/OPS_EXT_RedRose"
$ListTitle="VDR Data"

# Global Variables
$sCSOMPath = “C:\Users\prabhvx\OneDrive - Murphy Oil\SPLibs\” # Path to CSOM DLLs
$sTenantAdminUrl = “https://murphyoil-admin.sharepoint.com” # SharePoint Admin Site Collection
$sAdminUserName = “vinay_prabhugaonkar@contractor.murphyoilcorp.com” # Tenant Administrator username
$sAdminPassword = “Jun19@M0C#” | ConvertTo-SecureString -AsPlainText -Force # Tenant Administrator password

# Adding the Client OM Assemblies
$sCSOMRuntimePath = $sCSOMPath + “Microsoft.SharePoint.Client.Runtime.dll”
$sCSOMPath = $sCSOMPath + “Microsoft.SharePoint.Client.dll”

#Load SharePoint CSOM Assemblies

Add-Type -Path $sCSOMPath
Add-Type -Path $sCSOMRuntimePath

$ErrorActionPreference = "Continue"

Restore-SPOListAllItemsInheritance -Username $sAdminUserName -Url $siteUrl -AdminPassword $sAdminPassword -ListTitle $ListTitle

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow
