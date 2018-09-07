## SharePoint DLL 
Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
Start-SPAssignment –Global

$webApplicationURL = "http://thegrid.agl.com.au"
$out = "E:\thegridlistsize.csv"

$webApp = Get-SPWebApplication $webApplicationURL

if($webApp -ne $null)
{
	#"Web Application ; " + $webApp.Name | Out-File $out -Append

	foreach($siteColl in $webApp.Sites)
	{
    		if($siteColl -ne $null)
    		{
			#"Site Collection ; " + $siteColl.Url | Out-File $out -Append

			foreach($subWeb in $siteColl.AllWebs)
			{
				if($subWeb -ne $null)
				{
					#Print each Subsite
					#Write-Host $subWeb.Url
					#"Subsite ; " + $ + " - " + $subWeb.Url | Out-File $out -append
					foreach($list in $subweb.lists)
					{
                       $sizeofallitems=0
						foreach($item in $list.Items)
                        {
                            $sizeitem=($item.file).length
                            $sizeofallitems+=$sizeitem
                        }
                        #size in MB
                        $sizeoflist=($sizeofallitems/1024)/1024 
                        #size in KB
                        $sizeoflistkb=$sizeofallitems/1024 
                        
                       "$($siteColl.Name) `t $($siteColl.Url) `t $($subWeb.Name) `t $($subWeb.Url) `t $($list.Title) `t $($list.ItemCount) `t $($list.BaseTemplate) `t $($sizeoflist.ToString()) `t $($sizeoflistkb.ToString())" | Out-File $out -Append
                        
					}
					
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
