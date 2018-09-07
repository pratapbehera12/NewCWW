Add-PSSnapin "Microsoft.SharePoint.PowerShell"

$webapp = "http://portal"
$webappurl = get-spwebapplication $webapp


set-variable -option constant -name out -value "E:\Scripts\Sp_Migration\Printportal.csv"

foreach($spSite in $webappurl.Sites){



 
if($spSite -ne $null)
{
   "Site Collection : " + $spSite.Url | Out-File $out -Append
   foreach($subWeb in $spSite.AllWebs)
   {
      if($subWeb -ne $null)
      {
         #Print each Subsite
         #Write-Host $subWeb.Url
         "Subsite : " + $subWeb.Name + " - " + $subWeb.Url | Out-File $out -append
 
         $spListColl = $subweb.Lists
         foreach($eachList in $spListColl)
         {
            if($eachList.Title -eq "Pages")
            {
               $PagesUrl = $subweb.Url + "/"
               foreach($eachPage in $eachList.Items)
               {
                  #"Pages : " + $eachPage["Title"] + " - " + $PagesUrl + $eachPage.Url | Out-File $out -append
                  $PagesUrl + $eachPage.Url | Out-File $out -append
               }
            }
         }
         $subWeb.Dispose()
      }
      else
      {
         Echo $subWeb "does not exist"
      }
   }
   $spSite.Dispose()
}
else
{
   Echo $siteURL "does not exist, check the site collection url"
}

}
Echo Finish
