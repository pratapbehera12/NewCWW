Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
Start-SPAssignment –Global
 $csv="E:/Shaoli/griddetails.csv "
$webApplicationURL =  "	http://grid.agl.com.au"  
  $listcount =0
$webApp = Get-SPWebApplication $webApplicationURL  
   $timerjobs = Get-SPTimerJob

  $Results =$timerjobs.name 
#Write-Host "Web Application : " + $webApp.Name  
foreach($siteColl in $webApp.Sites)  
{  
if ($siteColl.AllowContentTypes -eq $true)
                                   {                           
                                           foreach ($sccontenttype in $siteColl.ContentTypes)
                                              {
                                                 $sitcolcontenttype=$sccontenttype.Name
                                                 }
                                }
 foreach($Subsites in $siteColl.AllWebs)
        
        {
        foreach($web in $Subsites)
            {
           # $groups = $web.RootWeb.sitegroups
                # foreach ($grp in $groups) 
                 #{
                # $grps= $grp.Name
                # }
                 
        if ([Microsoft.SharePoint.Publishing.PublishingWeb]::IsPublishingWeb($web)) 
               {
            $pWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
            $pages = $pWeb.PagesList
            foreach ($item in $pages.Items) 
                             {
                                      $manager = $item.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);
                                         $wps = $manager.webparts
                                       foreach($webpart in $wps)
                                          {
                                               $webPartName = $webpart.GetType().ToString()
                                               $webparttitle=  $webPartName.Title
                                           }   

                          
                            }
                      }
        
        else
            {
            $pages = $null
             $pages = $web.Lists[“Site Pages”]            
               if ($pages) 
                           {                
                                   foreach ($item in $pages.Items) 
                                   {
                                       $manager = $item.file.GetLimitedWebPartManager([System.Web.UI.WebControls.Webparts.PersonalizationScope]::Shared);
                                        $wps = $manager.webparts
                                       foreach($webpart in $wps)
                                    {
                                        $webPartName = $webpart.GetType().ToString()
                                        $webparttitle=  $webPartName.Title
                                    } 
                                 }
                             }
                 }         
         
                 
          
            
              
                 
                 
                 $lists = $web.Lists
                  foreach($list in $lists)
                  {
                  
                              foreach($item in $list.items) {
                           $listSize += ($item.file).length
                                  }
                      $filesize = [string][Math]::Round(($listSize/1KB),2)
                
                             $listitem=$list.ItemCount
                              if ($list.AllowContentTypes -eq $true)
                                   {                           
                                           foreach ($contenttype in $list.ContentTypes)
                                              {
                                                 $contenttypes=$contenttype.Name
                                                   foreach($listassociation in $list.WorkflowAssociations)
                                                  {
                                                   $associations += $($listassociation.Name)
                                                   $listcount +=1
                                                     }
        
        "$($siteColl.Url) `t $($Results.Title) `t $($sitcolcontenttype.Title)`t $($Subsites.Title) `t $($Subsites.Url)`t $($list.Title) `t $($list.DefaultViewUrl) `t $( $listitem)`t $( $filesize) `t $( $list.BaseType) `t $($contenttypes) `t $($associations)`t $( $listcount) `t $($list.LastItemModifiedDate) `t $($pages) `t $($webPartName) `t $( $webparttitle)"|Out-File $csv -Append
        
        
                                                    }  
                                              }


                       }  
}
}
}

