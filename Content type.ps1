Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
Start-SPAssignment –Global

$webapp = get-spwebapplication http://my.grid.agl.com.au
$logtext = "E:\SP_Migration\contenttypemygrid.csv"

foreach($sites in $webapp.Sites){


$web = $sites.RootWeb

foreach ($ctype in $web.ContentTypes) {

$content = $ctype.Name 
$contentresource = $ctype.NameResource
 "$($content) `t $($contentresource)" | Out-File $logtext -Append


}



}