## SharePoint DLL 
Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
Start-SPAssignment –Global

Get-SPFeature -Limit ALL | Where-Object {$_.Scope -eq "SITE"} | Export-csv "E:\SPfeatures.csv"