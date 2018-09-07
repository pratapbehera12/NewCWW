## SharePoint DLL 
Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
Start-SPAssignment –Global
    $FileLocation = "E:\Souvick\SharePointGroupMyGrid.csv"
    

    #Write CSV- TAB Separated File) Header
   # "URL `t SiteOwnerGroup `t SiteOwnerDisplayName `t OwnerofGroup" | Out-file $FileLocation 

    #Get the specified Web Application
    $web1="http://my.grid.agl.com.au"
    $WebApp= Get-SPWebApplication $web1
  # $x="";

    #Loop through all site collections in Web Application
    foreach($Site in $webapp.sites)
    {
        #Loop throuh all Sub Sites
        foreach($Web in $Site.AllWebs)
        {
        # $count=1
            foreach($Group in $Web.Groups)
            { 
                                         #Send the Data to Log file
                    "$($WebApp.Name) `t $($Site.Url) `t $($web.url) `t $($Group.Name) `t $($Group.Roles) `t $($Group.Owner) `t $($Group.Owner.DisplayName) " | Out-file $FileLocation  -Append
                
            
              
            }
            
        }
    }


 

Remove-PSSnapin Microsoft.SharePoint.PowerShell 
 
Write-Output "" 
Write-Output "Script Execution finished" 