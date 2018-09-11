[IO.Directory]::SetCurrentDirectory((Convert-Path (Get-Location -PSProvider FileSystem)))

# check if we are in the same location as the nwadmin.exe
if(Test-Path("C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\BIN\nwadmin.exe"))
{ 
  # find all the workflows and store them in a variable
  $foundworkflows = nwadmin.exe -o FindWorkflows

  foreach($line in $foundworkflows)
  {
    if($line.StartsWith("Active at "))
    {
      # get the site url
      $site = $line.Replace("Active at ","");
    }
    if($line.StartsWith("-- "))
    { 
      # get the list name
      $list = $line.Replace("-- ","");
    }
    if($line.StartsWith("---- "))
    {
      # get the workflow name
      $workflowname = $line.Replace("---- ","");

      $message = "{0} - {1} - {2}" -f $site,$list,$workflowname;
      echo $message;
    }
  }
}
else
{
  echo "NWAdmin doesn't exist.  Change directory to where NWAdmin.exe lives.";
}
