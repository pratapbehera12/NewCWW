Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
Start-SPAssignment –Global

$csv="E:/Souvick/thegridwebpart.csv"
$webApplicationURL =  "http://thegrid.agl.com.au"  


$webApp = Get-SPWebApplication $webApplicationURL  

Write-Host "Web Application : " + $webApp.Name  
$webparttitle = @()

foreach($siteColl in $webApp.Sites)  
{  

	foreach($Subsites in $siteColl.AllWebs)
    {
       	foreach($web in $Subsites)
         {
	        	
                 
	        if ([Microsoft.SharePoint.Publishing.PublishingWeb]::IsPublishingWeb($web)) 
            {
         		$pWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
		        $list = $Web.Lists["Pages"]
		        foreach ($item in $list.Items) 
                        {
                        	$manager = $item.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);
                                $wps = $manager.WebParts
                                foreach($webpart in $wps)
                                {
                                	$webPartName = $webpart.GetType().ToString()
                                        $webparttitle=  $webPart.Title
                                        "$($siteColl.Url) `t $($Subsites.Title) `t $($Subsites.Url)`t $($list) `t $($item.Url) `t $($webPartName) `t $( $webparttitle)"|Out-File $csv -Append

                                }   

                          
                        }

             }
        
	        else
        	{
		        $list = $null
	                $list = $web.Lists[“Site Pages”]            
           		if ($list) 
                        {                
                        	foreach ($item in $list.Items) 
                                {
                                	$manager = $item.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);
                                        $wps = $manager.webparts
                                        foreach($webpart in $wps)
                                    	{
                                        	$webPartName = $webpart.GetType().ToString()
	                                        $webparttitle=  $webPart.Title
                                         "$($siteColl.Url) `t $($Subsites.Title) `t $($Subsites.Url)`t $($list) `t $($item.Url) `t $($webPartName) `t $( $webparttitle)"|Out-File $csv -Append
        	                            } 
                                 }
                                 
                        }
                }         
         
                 
              
        
        #"$($siteColl.Url) `t $($Subsites.Title) `t $($Subsites.Url)`t $($pages) `t $($webPartName) `t $( $webparttitle)"|Out-File $csv -Append
        
        }
        }  
}

  

