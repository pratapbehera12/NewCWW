Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
Start-SPAssignment –Global
 $csv="E:/Shaoli/gridTimerdetails_07-19-2018.csv "
$webApplicationURL =  "http://grid.agl.com.au"  
  $listcount =0
$webApp = Get-SPWebApplication $webApplicationURL  
   $timerjobs = Get-SPTimerJob

  $Results =$timerjobs.name | Out-File $csv -Append