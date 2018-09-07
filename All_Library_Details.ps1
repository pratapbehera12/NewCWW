## SharePoint DLL 
Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
Start-SPAssignment –Global

$OutputReport = "E:\SP_Migration\Alllibrary_portal_18-07-2018.csv"  
$webapplication = "http://portal"

$webapp = get-SPwebapplication $webapplication 

foreach($sites in $webapp.Sites){


$wc = $sites.AllWebs
foreach($w in $wc)
{
foreach($l in $w.Lists){
 
        $count = $l.ItemCount
       
          "$($sites.url) `t $($w.Url) `t $($w)`t $($l) `t $($l.BaseTemplate) `t $($count) `t " | Out-File $OutputReport -Append  


}
}
}