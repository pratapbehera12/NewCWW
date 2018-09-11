Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
Start-SPAssignment –Global
  $out="E:/testEventReceiver1.csv "
$webApplicationURL =  "http://grid.agl.com.au"  
  
$webApp = Get-SPWebApplication $webApplicationURL  
  foreach($siteColl in $webApp.Sites)
           {
                if($siteColl.url -eq "http://grid.agl.com.au/sites/Spaces02"){
                 foreach($Subsites in $siteColl.AllWebs)
                 {
                            
                     if($Subsites.Url -eq "http://grid.agl.com.au/sites/Spaces02/MktDev"){
                                foreach($list in $Subsites.Lists )
                                          
                                          {
                                            if($list.Title -eq "Drop Off Library")
                                            {
                                            
                                                  $EventReceiver=$list.EventReceivers.Class

                                                  if($EventReceiver -ne ""){

                                                  foreach($event in $EventReceiver){
                                                
                                               
                                             
                                                "$($siteColl.Url)  `t $($Subsites.Url) `t $( $SiteEvRec)`t $($list.Title)`t $($event)"|  Out-File $out -Append

                                                }
                                                }
                                                }
                    }
                    }
                    }
                    }
                    }
          
          

          