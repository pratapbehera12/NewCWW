## SharePoint DLL 
Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
Start-SPAssignment –Global

$webApplicationURL = "http://my.grid.agl.com.au"
$out = "E:\SP_Migration\PrintAllSitesthegrid.csv"

$webApp = Get-SPWebApplication $webApplicationURL

if($webApp -ne $null)
{
"Web Application ; " + $webApp.Name | Out-File $out -Append

foreach($siteColl in $webApp.Sites)
{
    if($siteColl -ne $null)
    {
"Site Collection ; " + $siteColl.Url | Out-File $out -Append

foreach($subWeb in $siteColl.AllWebs)
{
if($subWeb -ne $null)
{
#Print each Subsite
#Write-Host $subWeb.Url
"Subsite ; " + $subWeb.Name + " - " + $subWeb.Url | Out-File $out -append
                  
$subWeb.Dispose()
}
else
{
Echo $subWeb "does not exist"
}
}
$siteColl.Dispose()
}
else
{
Echo $siteColl "does not exist"
}
}
}
else
{
Echo $webApplicationURL "does not exist, check the WebApplication name"
}
Stop-SPAssignment -Global
Remove-PsSnapin Microsoft.SharePoint.PowerShell

Echo Finish