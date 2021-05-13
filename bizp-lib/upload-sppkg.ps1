Write-Host "========================================================================" -ForegroundColor Green
Write-Host "Uploading SPPKG file" -ForegroundColor Yellow
$jsonObj = Get-Content './config/package-solution.json' | Out-String | ConvertFrom-Json
Write-Host " File Path : ./sharepoint/$($jsonObj.paths.zippedPackage)" -ForegroundColor Yellow
Write-Host " Solution Id : $($jsonObj.solution.id)" -ForegroundColor Yellow


#$creds = Get-Credential -UserName "ggoyal@o365code.onmicrosoft.com" -Message "Upload SPPKG File."
#$siteurl="https://o365code.sharepoint.com/sites/AppCat"

$siteurl="https://m365x039710.sharepoint.com/sites/AppCat"

Write-Host "Connecting to App Catalog: $($siteurl)" -ForegroundColor Yellow
$creds = Get-Credential -UserName "ggoyal@M365x039710.onmicrosoft.com" -Message "Upload SPPKG File."
Connect-PnPOnline -Url $siteurl -Credentials $creds
Write-Host "Connected to App Catalog: $($siteurl)" -ForegroundColor Green
#Add-PnPApp -Path ./sharepoint/solution/bizportal-wiki-center.sppkg -Overwrite -Publish
#https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/add-pnpapp?view=sharepoint-ps

Write-Host "Adding App to App Catalog: $($siteurl)" -ForegroundColor Yellow
Add-PnPApp -Path "./sharepoint/$($jsonObj.paths.zippedPackage)" -Overwrite -Scope Tenant -Publish
Write-Host "Added App to App Catalog: $($siteurl)" -ForegroundColor Green

Write-Host "Publising App" -ForegroundColor Yellow
#https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/publish-pnpapp?view=sharepoint-ps
Publish-PnPApp -Identity $jsonObj.solution.id -SkipFeatureDeployment -Scope Tenant
Write-Host "Publised App" -ForegroundColor Green
#And install at any particular site:
#Install-PnPApp -Identity <app id> -Scope Tenant

Write-Host "Completed SPPKG file related operation" -ForegroundColor Green
