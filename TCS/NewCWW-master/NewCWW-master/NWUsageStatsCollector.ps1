<#
.SYNOPSIS
Nintex Workflow Usage Statistics

.DESCRIPTION
This script is used to collect usage statistics from all the Nintex Workflow databases. By executing this script, files in CSV or XML containing the statistics will be generated and saved into a zip file.

.PARAMETER fileName
Name of the output file to be saved. This is an optional parameter. Default value is NWUsageStats i.e. name of the zip file generated with the statistics will be NWUsageStats

.PARAMETER outputFormat
Format of the output file. It should be either CSV or XML. This is an optional parameter and if it is not included, then files in both the formats will be generated and saved into zip file

.PARAMETER excludeFields
Fields to be excluded in SQL query. Currently only "WorkflowName" field is supported, i.e. you can only exclude name of the workflows from the usage statistics

.EXAMPLE

Save usage statistics result into a zip called NWUsageStats with date and time appended in yyyyMMdd-HHmmss format.

.\NWUsageStatsCollector.ps1 

.EXAMPLE

Save the zip file with user defined name and exclude workflow name in the output statistics file.

.\NWUsageStatsCollector.ps1 -fileName "Result" -excludeFields "WorkflowName"

.EXAMPLE

Save the zip file with user defined name, exclude workflow name from the output statistics and only include CSV format in zip file.

NWUsageStatsCollector.ps1 -fileName "Result" –outputFormat CSV –excludeFields "WorkflowName"

.NOTES

You need to run this script from SharePoint Management Shell

#>

param(
 [string]$fileName = "NWUsageStats",  [ValidateSet("xml","csv", "both")] [String] $outputFormat = "both", [string] $excludeFields)


if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null)
{
	Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
}

[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow")
[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.Administration")

$excludeFieldsArray = $excludeFields -split "," | ForEach-Object { $_.Trim().ToLower() }

function IsFileLocked(
    [string] $path)
{
    If ([string]::IsNullOrEmpty($path) -eq $true)
    {
        Throw "The path must be specified."
    }
    
    [bool] $fileExists = Test-Path $path
    
    If ($fileExists -eq $false)
    {
        Throw "File does not exist (" + $path + ")"
    }
    
    [bool] $isFileLocked = $true

    $file = $null
    
    Try
    {
        $file = [IO.File]::Open(
            $path,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            [IO.FileShare]::None)
            
        $isFileLocked = $false
    }
    Catch [IO.IOException]
    {
        If ($_.Exception.Message.EndsWith(
            "it is being used by another process.") -eq $false)
        {
            Throw $_.Exception
        }
    }
    Finally
    {
        If ($file -ne $null)
        {
            $file.Close()
        }
    }
    
    return $isFileLocked
}

function Replace-UnicodeCharacters {
param([string]$inputString, [char]$replaceChar)
  $value = [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($inputString))
  $value.replace('?',$replaceChar)
}

function Save-Result([System.Object]$result, $fileName, $outputFormat)
{
	if([string]::IsNullOrEmpty($fileName))
	{
		$fileName = "NWUsageStats"
	}
		
	$directory = [io.path]::GetDirectoryName($fileName)
	if([string]::IsNullOrEmpty($directory))
	{
		$directory = (Convert-Path ".")
	}
	
	$directory = $directory.TrimEnd('\')
	
	if(!(Test-Path -Path $directory )){
   		New-Item -ItemType directory -Path $directory
	}
	
	$filename = [io.path]::GetFileNameWithoutExtension($fileName)
	
	$skipXML = $outputFormat.ToLower() -eq "csv"
	$skipCSV = $outputFormat.ToLower() -eq "xml"
	
	$savedFiles = @()
	
	if($skipCSV -eq $false)
	{
		$filePath =  "$directory\$filename-$(get-date -f yyyyMMdd-HHmmss).csv"
		
		foreach ($nwUsageStats in $result.NWUsageStatsSet)
		{
			$dbName = $nwUsageStats.ContentDBName
			$reportDate = $nwUsageStats.ReportDate.ToString("yyyyMMdd-HHmmss")
			
			if($nwUsageStats.OverviewStatistics -ne $null) { 
				$filePath =  "$directory\" + (Replace-UnicodeCharacters "$filename-$dbName-OverviewStatistics-$reportDate.csv" '_')
				$nwUsageStats.OverviewStatistics | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
				$savedFiles += $filePath
			}
			if($nwUsageStats.WorkflowStatistics -ne $null) { 
				$filePath =  "$directory\" + (Replace-UnicodeCharacters "$filename-$dbName-WorkflowStatistics-$reportDate.csv" '_')
				$nwUsageStats.WorkflowStatistics | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
				$savedFiles += $filePath
			}
			if($nwUsageStats.ActionStatistics -ne $null) { 
				$filePath =  "$directory\" + (Replace-UnicodeCharacters "$filename-$dbName-ActionStatistics-$reportDate.csv" '_')
				$nwUsageStats.ActionStatistics  | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
				$savedFiles += $filePath
			}
		}
		
		$configDBName = $result.ConfigDBName
		$resultID = $result.ResultID
		$filePath =  "$directory\" + (Replace-UnicodeCharacters "$filename-$configDBName-Activities-$resultID.csv" '_') 
		$result.Activities | Export-Csv -Path $filePath -NoTypeInformation -Encoding Utf8
		$savedFiles += $filePath
	}
	
	if($skipXML -eq $false)
	{
		$filePath =  "$directory\" + (Replace-UnicodeCharacters "$filename-$(get-date -f yyyyMMdd-HHmmss).xml" '_')
		$result | Export-Clixml $filePath -Encoding Utf8
		$savedFiles += $filePath
	}
	
	$zipFilePath = "$directory\$filename-$(get-date -f yyyyMMdd-HHmmss).zip"
	Add-Zip -fileList $savedFiles -zipfilename $zipFilePath
	
	Write-Host "Results saved to $zipFilePath."
}

function Add-Zip
{
    param($fileList, [string]$zipfilename)

    if(-not (test-path($zipfilename)))
    {
        set-content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
        (dir $zipfilename).IsReadOnly = $false  
    }
	
    $shellApplication = new-object -com shell.application
    $zipPackage = $shellApplication.NameSpace($zipfilename)
	
    foreach($f in $fileList) 
    { 
		$file = Get-ChildItem $f 
        $zipPackage.CopyHere($file.FullName)
		
		while($zipPackage.Items().Item($file.name) -eq $null){
        	Start-sleep -seconds 1
    	}
		
		Remove-Item $f
    } 
}

function Test-NWInstalled {
	$assembly = [System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow")
	if($assembly -eq $null)
	{
		Write-Host "Nintex Workflow is not installed. Process will be exited."
		return $false
	}
	else
	{
		return $true
	}
}

function GetProductVersion
{
    process
    {
        trap [Management.Automation.RuntimeException]
        {
            return $null
        }

        Invoke-Expression "[Nintex.Workflow.Licensing.License]::VersionNumberInfo"
    }
}

function Test-DBConnection {
	param ($connectionString)
	try {
		$conn = New-DBConnection $connectionString
		if($conn.State -eq "Closed") { 
			Write-Host "Failed to establish a connection to Nintex Workflow configuration database. Check your connection string."
			return $false
		}
		else { Remove-DBConnection $conn }
		
		return $true
	  }
    catch {
      Write-Host "Failed to establish a connection to Nintex Workflow configuration database. Check your connection string."
	  return $false
    }
}

function New-DBConnection {
    param ($connectionString)
    if (test-path variable:\conn) {
        $conn.close()
    } else {
        $conn = new-object System.Data.SqlClient.SQLConnection($connectionString)
    }
    $conn.Open()
    $conn
}

function Remove-DBConnection {
    param ($connection)
    $connection.close()
    $connection = $null
}

function Execute-Scalar {
	param ($query, $connection)
        
        $cmd = new-object System.Data.SqlClient.SqlCommand
		$cmd.CommandText = $query
		$cmd.CommandTimeout = 7200
		$cmd.Connection = $connection
        $cmd.ExecuteScalar();
		
        
}

function Execute-Reader {
	param($query, $connection)
	
        $cmd = new-object System.Data.SqlClient.SqlCommand
		$cmd.CommandText = $query
		$cmd.CommandTimeout = 7200
		$cmd.Connection = $connection
        $dr = $cmd.ExecuteReader()
		
		,$dr
}

function Query-Rows {
	param ($query, $connectionString)
	
	$connection = New-DBConnection $connectionString
	
	$reader =  Execute-Reader $query $connection
	$rows = @()
	$counter = $reader.FieldCount
	while($reader.Read())
	{
		$row = New-Object System.Object
		for ($i = 0; $i -lt $counter; $i++)
		{
			$row | Add-Member -MemberType NoteProperty -Name $reader.GetName($i) -Value $reader.GetValue($i)
		}
		$rows += $row
	}
	$reader.Close()
	
	Remove-DBConnection $connection
	
	$rows
}

function Get-ReportDate {
	param ([Nintex.Workflow.Administration.ContentDatabase] $db)
	
	$query = "Select getUTCdate() as ReportDate"
	
	$connection = New-DBConnection $db.SQLConnectionString.ToString()
	
	Execute-Scalar $query $connection
	
	Remove-DBConnection $connection
}

<#
  High-level statistics
#>
function Get-OverviewStatistics {
	param ([Nintex.Workflow.Administration.ContentDatabase] $db)
	
	$query = "Select
						Stats.Yr,
						Stats.Mth,
						Stats.UniqueSiteCollections,
						Stats.UniqueSites,
						Stats.UniqueInitiators,
						Stats.UniqueWorkflows,
						Stats.CompletedWorkflows,
						UA.UniqueApproverEmailAddresses,
						UA.UniqueApproverUsernames,
						UA.TotalApprovers
					from
					(
					--stats
					Select 
						Year(StartTime) as Yr,
						MONTH(StartTime) as Mth,
						count(distinct SiteID) as UniqueSiteCollections,
						count(distinct WebID) as UniqueSites,
						count(distinct WorkflowID) as UniqueWorkflows,
						count(WorkflowInstanceID) as CompletedWorkflows,
						count(distinct WorkflowInitiator) as UniqueInitiators
					from 
						WorkflowInstance with (NOLOCK)
					where 
						StartTime > DATEADD(m,-13, getUTCdate())
						and State = 4
					group by 
						Year(StartTime), Month(StartTime)
						) Stats left join
					(
					--Unique approvers
					select
						Year(EntryTime) as Yr,
						MONTH(EntryTime) as Mth,
						count(distinct Username) as UniqueApproverUsernames,
						count(distinct EmailAddress) as UniqueApproverEmailAddresses,
						count(Username) as TotalApprovers
					from 
						HumanWorkflowApprovers with (NOLOCK)
					where 
						EntryTime > DATEADD(m,-13, getUTCdate())
					group by 
						Year(EntryTime), Month(EntryTime)
					) UA
					on Stats.Yr = UA.Yr and Stats.Mth = UA.Mth
					order by 
						1,2"
	
	Query-Rows $query $db.SQLConnectionString.ToString()
}

<#
  Workflows that have completed, Max actions, Min actions, Avg actions, last run date, # unique approvers, #unique imitators
#>
function Get-WorkflowStatstics {
	param ([Nintex.Workflow.Administration.ContentDatabase] $db, $day)
	
	if($day -ne $null)
	{
		$timeStampFilter = " and TimeStamp > DATEADD(d,-$day, getUTCdate()) "
		$startTimeFilter = " and I.StartTime > DATEADD(d,-$day, getUTCdate()) "
	}
	else
	{
		$timeStampFilter = " "
		$startTimeFilter = " "
	}
	
	$includeWorklowNameColumn = ""
	if(!$excludeFieldsArray.Contains("workflowname"))
	{
		$includeWorklowNameColumn = "I.WorkflowName,"
	}
	
	$query = "Select
						I.WorkflowID,
						$includeWorklowNameColumn
						Max(P.ActionExecutions) as MaxActionExecutions,
						Min(P.ActionExecutions) as MinActionExecutions,
						Avg(P.ActionExecutions) as AvgActionExecutions,
						max(P.SequenceID) as MaxDesignerActions,
						count(distinct WorkflowInitiator) as UniqueInitiators,
						Max(I.StartTime) as LastRunDate,
						Max(Datediff(ss, I.StartTime, P.LastActivityTime)) as MaxRunDurationSeconds,
						Min(Datediff(ss, I.StartTime, P.LastActivityTime)) as MinRunDurationSeconds,
						count(I.InstanceID) as TotalInstancesRun
					from
					WorkflowInstance I with (NOLOCK) inner join
					(select 
						InstanceID as InstanceID,
						count(WorkflowProgressID) as ActionExecutions,
						max(SequenceID) as SequenceID,
						Max(TimeStamp) as LastActivityTime
					from 
					WorkflowProgress with (NOLOCK)
					where
						ActivityComplete = 1 $timeStampFilter 
						group by
						InstanceID
					) P
					on I.InstanceID = P.InstanceID
					where 
						I.State = 4 $startTimeFilter
					 group by
						I.WorkflowID,
						I.WorkflowName
					order by 2"
					
	Query-Rows $query $db.SQLConnectionString.ToString()
}

<#
  Actions (TypeName), how many times executed, number of workflow instaces, # appeared in unique workflows
#>
function Get-ActionStatstics {
	param ([Nintex.Workflow.Administration.ContentDatabase] $db, $day)
	
	if($day -ne $null)
	{
		$timeStampFilter = " AND TimeStamp > DATEADD(d,-$day, getUTCdate()) "
		$startTimeFilter = " WHERE StartTime > DATEADD(d,-$day, getUTCdate()) "
	}
	else
	{
		$timeStampFilter = " "
		$startTimeFilter = " "
	}
	
	$query = "Select 
				Actions.ActivityID,
				Actions.TotalExecutions,
				Actions.TotalWorkflowInstances,
				UsedIn.NumberOfWorkflowsUsedIn
				 from
				(
				--Actions
				Select
					ActivityID,
					count(WorkflowProgressID) as TotalExecutions,
					count(distinct InstanceID) as TotalWorkflowInstances
				from WorkflowProgress with (NOLOCK)
				where
					ActivityComplete = 1
					$timeStampFilter
				group by ActivityID) Actions inner join
				(Select
					ActivityID,
					count(distinct WorkflowID) as NumberOfWorkflowsUsedIn
				from WorkflowProgress with (NOLOCK) inner join WorkflowInstance with (NOLOCK)
				ON WorkflowProgress.InstanceID = WorkflowInstance.InstanceID
				$startTimeFilter
				group by ActivityID) UsedIn
				on Actions.ActivityID = UsedIn.ActivityID"
	
	Query-Rows $query $db.SQLConnectionString.ToString()
}

<#
	Get the activities list  from config DB
#>
function Get-ActivityList {
	param([Nintex.Workflow.Administration.ConfigurationDatabase] $db)
	
	$query = "IF EXISTS (SELECT * FROM sys.objects 
						WHERE object_id = OBJECT_ID(N'[dbo].[Activities]') AND type in (N'U'))
						BEGIN
						Select ActivityID, ActivityName, ActivityAssembly, ActivityType from Activities
					END"
					
	Query-Rows $query $db.SQLConnectionString.ToString()
}

<#
	Check Permission to Access Farm via PowerShell
#>
function Test-FarmPermission {
        
    try
    {
	    Get-SPFarm | out-null 
    }
    catch
    {
	    write-error $($_.Exception.Message)		
        exit	
    }
}

Write-Host "======Nintex Workflow Usage Statistics======"
Write-Host


$isNWINstalled = Test-NWInstalled -eq $true

if($isNWINstalled)
{
	$productVersion = GetProductVersion
	$productName = "Nintex Workflow "
	if($productVersion.StartsWith("1"))
	{
		$productName += "2007 "
	}
	elseif($productVersion.StartsWith("2"))
	{
		$productName += "2010 "
        Test-FarmPermission
	}
	elseif($productVersion.StartsWith("3"))
	{
		$productName += "2013 "
        Test-FarmPermission
	}
	
	Write-Host "Installed Nintex Workflow version: $productName$productVersion"
		
	$configDB = [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase()
	$isDBConnectable = Test-DBConnection $configDB.SQLConnectionString.ToString()
	if($isDBConnectable)
	{
		$configDBName = $configDB.SQLConnectionString.InitialCatalog
		$configDBVersion = [Nintex.Workflow.Administration.ConfigurationDatabase]::DatabaseVersion
		$contentDatabases = $configDB.ContentDatabases
		
		Write-Host "Nintex Workflow Configuration Database version: $configDBName ($configDBVersion)"
		Write-Host "Number of content databases: $($contentDatabases.Count)"
		
		$NWUsageStatsSet = @()

		foreach ($database in $contentDatabases)
		{
			Write-Host
			$dbName = $database.SqlConnectionString.InitialCatalog
			
			Write-Host "Collecting usage statistics from content database $dbName..."
			
			$reportDate = Get-ReportDate $database	
			
			Write-Host -NoNewline ">>> Retrieving overview statistics..."
			$t = Measure-Command { $overviewStats = Get-OverviewStatistics $database }
			Write-Host "Query duration $( '{0:f2}' -f $t.TotalSeconds ) seconds."
			
			Write-Host -NoNewline ">>> Retrieving workflow statistics..."
			$t = Measure-Command { $workflowStats = Get-WorkflowStatstics $database }
			Write-Host "Query duration $( '{0:f2}' -f $t.TotalSeconds ) seconds."
			
			Write-Host -NoNewline ">>> Retrieving action statistics..."
			$t = Measure-Command { $actionStats = Get-ActionStatstics $database }
			Write-Host "Query duration $( '{0:f2}' -f $t.TotalSeconds ) seconds."
			
			$NWUsageStats = New-Object System.Object
			$NWUsageStats | Add-Member -MemberType NoteProperty -Name "ContentDBName" -Value $dbName
			$NWUsageStats | Add-Member -MemberType NoteProperty -Name "ReportDate" -Value $reportDate
			$NWUsageStats | Add-Member -MemberType NoteProperty -Name "OverviewStatistics" -Value $overviewStats
			$NWUsageStats | Add-Member -MemberType NoteProperty -Name "WorkflowStatistics" -Value $workflowStats
			$NWUsageStats | Add-Member -MemberType NoteProperty -Name "ActionStatistics" -Debug $actionStats
			$NWUsageStatsSet += $NWUsageStats
			
			Write-Host "DONE Collecting usage statistics from content database $dbName..."
		}
		Write-Host
		
		Write-Host "Collecting activities list from configuration database $configDBName..."
		Write-Host -NoNewline ">>> Retrieving activities list..."
		$t = Measure-Command { $Activities = Get-ActivityList $configDB }
		Write-Host "Query duration $( '{0:f2}' -f $t.TotalSeconds ) seconds."

		$NWUsageStatsResult = New-Object System.Object
		$NWUsageStatsResult | Add-Member -MemberType NoteProperty -Name "ResultID" -Value ([guid]::NewGuid())
		$NWUsageStatsResult | Add-Member -MemberType NoteProperty -Name "ConfigDBName" -Value $configDBName
		$NWUsageStatsResult | Add-Member -MemberType NoteProperty -Name "NWUsageStatsSet" -Value $NWUsageStatsSet
		$NWUsageStatsResult | Add-Member -MemberType NoteProperty -Name "Activities" -Value $Activities
		
		Write-Host
		Save-Result -result $NWUsageStatsResult -fileName $fileName -outputFormat $outputFormat
	}
}

Write-Host
Write-Host "======Finished======"
exit
# SIG # Begin signature block
# MIIa7wYJKoZIhvcNAQcCoIIa4DCCGtwCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUZ7gusZTN0G8NGzft7r+rPjiY
# 2qqgghZJMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
# AQUFADCBizELMAkGA1UEBhMCWkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIG
# A1UEBxMLRHVyYmFudmlsbGUxDzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhh
# d3RlIENlcnRpZmljYXRpb24xHzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcg
# Q0EwHhcNMTIxMjIxMDAwMDAwWhcNMjAxMjMwMjM1OTU5WjBeMQswCQYDVQQGEwJV
# UzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFu
# dGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBALGss0lUS5ccEgrYJXmRIlcqb9y4JsRDc2vCvy5Q
# WvsUwnaOQwElQ7Sh4kX06Ld7w3TMIte0lAAC903tv7S3RCRrzV9FO9FEzkMScxeC
# i2m0K8uZHqxyGyZNcR+xMd37UWECU6aq9UksBXhFpS+JzueZ5/6M4lc/PcaS3Er4
# ezPkeQr78HWIQZz/xQNRmarXbJ+TaYdlKYOFwmAUxMjJOxTawIHwHw103pIiq8r3
# +3R8J+b3Sht/p8OeLa6K6qbmqicWfWH3mHERvOJQoUvlXfrlDqcsn6plINPYlujI
# fKVOSET/GeJEB5IL12iEgF1qeGRFzWBGflTBE3zFefHJwXECAwEAAaOB+jCB9zAd
# BgNVHQ4EFgQUX5r1blzMzHSa1N197z/b7EyALt0wMgYIKwYBBQUHAQEEJjAkMCIG
# CCsGAQUFBzABhhZodHRwOi8vb2NzcC50aGF3dGUuY29tMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwPwYDVR0fBDgwNjA0oDKgMIYuaHR0cDovL2NybC50aGF3dGUuY29tL1Ro
# YXd0ZVRpbWVzdGFtcGluZ0NBLmNybDATBgNVHSUEDDAKBggrBgEFBQcDCDAOBgNV
# HQ8BAf8EBAMCAQYwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0y
# MDQ4LTEwDQYJKoZIhvcNAQEFBQADgYEAAwmbj3nvf1kwqu9otfrjCR27T4IGXTdf
# plKfFo3qHJIJRG71betYfDDo+WmNI3MLEm9Hqa45EfgqsZuwGsOO61mWAK3ODE2y
# 0DGmCFwqevzieh1XTKhlGOl5QGIllm7HxzdqgyEIjkHq3dlXPx13SYcqFgZepjhq
# IhKjURmDfrYwggQgMIIDCKADAgECAhA0TtVXINXt7En0L8432yttMA0GCSqGSIb3
# DQEBBQUAMIGpMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMdGhhd3RlLCBJbmMuMSgw
# JgYDVQQLEx9DZXJ0aWZpY2F0aW9uIFNlcnZpY2VzIERpdmlzaW9uMTgwNgYDVQQL
# Ey8oYykgMjAwNiB0aGF3dGUsIEluYy4gLSBGb3IgYXV0aG9yaXplZCB1c2Ugb25s
# eTEfMB0GA1UEAxMWdGhhd3RlIFByaW1hcnkgUm9vdCBDQTAeFw0wNjExMTcwMDAw
# MDBaFw0zNjA3MTYyMzU5NTlaMIGpMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMdGhh
# d3RlLCBJbmMuMSgwJgYDVQQLEx9DZXJ0aWZpY2F0aW9uIFNlcnZpY2VzIERpdmlz
# aW9uMTgwNgYDVQQLEy8oYykgMjAwNiB0aGF3dGUsIEluYy4gLSBGb3IgYXV0aG9y
# aXplZCB1c2Ugb25seTEfMB0GA1UEAxMWdGhhd3RlIFByaW1hcnkgUm9vdCBDQTCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAKyg8PuAWdScx6TPnaFZcwkQ
# RQwNLG5o8WxbSGhJWTf8CzMZwnd/zBAtlTQc5utNCacc0rjJlzYCt4nUJF8GwMxE
# lJSNAmJv61rdEY0omlyEkBB6Db10Zi9qOKDi1VRE6x0Hnwe6b+7p/U4LKfU+hKAB
# 8Zyr+Bx+iaToodhxZQ2jUXvuvNIiYA25W53fuvxRWwuvmLLpLukE6GKH3ivI107B
# TGQe3c+HWLpKT8poBx0cnUrG1S+RzHxxchzFwGfrMv3JklyU2oXAm79TfSsJ9Iyd
# kR+XalLL3gk2pHfYe4dQRNU+bilp+zlJJh4JpYB7QC3r6CeFyf5h/X7mfJcd1Z0C
# AwEAAaNCMEAwDwYDVR0TAQH/BAUwAwEB/zAOBgNVHQ8BAf8EBAMCAQYwHQYDVR0O
# BBYEFHtbRc+vzst6/TGSGmq280brV0hQMA0GCSqGSIb3DQEBBQUAA4IBAQB5EcBL
# s5G2/PDpZ9QNbkW+VeiT0s4DP+3aJbAdV8seOnagTOxQduhkcgykqfG4i9bWh4S7
# MuVBEcB32bNgnesb1dFuRESppgHsVWIdd7hcjkhJfJw7VxGsrXM3ji94XJBoR9lg
# YOb8Bz0iIBfE9xbpxNhy+chzfN8WLxWpPv1qJ7ah61q6mB/V401kCp0TyGG69Tkc
# h7q4vXsif/b+rEB55awQbz2PG3l2i8Q3syEYhOU2AOtjIJm56f4zBLtByMEC+URj
# IJ6BzkLT1j8sdtNjnFndj6bhDqAuQfculUfPvP0z8/YLYX5+kSuBR8InMO6nEF03
# j1w5K+QE8HuNVoxoMIIEmTCCA4GgAwIBAgIQcaC3NpXdsa/COyuaGO5UyzANBgkq
# hkiG9w0BAQsFADCBqTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDHRoYXd0ZSwgSW5j
# LjEoMCYGA1UECxMfQ2VydGlmaWNhdGlvbiBTZXJ2aWNlcyBEaXZpc2lvbjE4MDYG
# A1UECxMvKGMpIDIwMDYgdGhhd3RlLCBJbmMuIC0gRm9yIGF1dGhvcml6ZWQgdXNl
# IG9ubHkxHzAdBgNVBAMTFnRoYXd0ZSBQcmltYXJ5IFJvb3QgQ0EwHhcNMTMxMjEw
# MDAwMDAwWhcNMjMxMjA5MjM1OTU5WjBMMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# dGhhd3RlLCBJbmMuMSYwJAYDVQQDEx10aGF3dGUgU0hBMjU2IENvZGUgU2lnbmlu
# ZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJtVAkwXBenQZsP8
# KK3TwP7v4Ol+1B72qhuRRv31Fu2YB1P6uocbfZ4fASerudJnyrcQJVP0476bkLjt
# I1xC72QlWOWIIhq+9ceu9b6KsRERkxoiqXRpwXS2aIengzD5ZPGx4zg+9NbB/BL+
# c1cXNVeK3VCNA/hmzcp2gxPI1w5xHeRjyboX+NG55IjSLCjIISANQbcL4i/CgOaI
# e1Nsw0RjgX9oR4wrKs9b9IxJYbpphf1rAHgFJmkTMIA4TvFaVcnFUNaqOIlHQ1z+
# TXOlScWTaf53lpqv84wOV7oz2Q7GQtMDd8S7Oa2R+fP3llw6ZKbtJ1fB6EDzU/K+
# KTT+X/kCAwEAAaOCARcwggETMC8GCCsGAQUFBwEBBCMwITAfBggrBgEFBQcwAYYT
# aHR0cDovL3QyLnN5bWNiLmNvbTASBgNVHRMBAf8ECDAGAQH/AgEAMDIGA1UdHwQr
# MCkwJ6AloCOGIWh0dHA6Ly90MS5zeW1jYi5jb20vVGhhd3RlUENBLmNybDAdBgNV
# HSUEFjAUBggrBgEFBQcDAgYIKwYBBQUHAwMwDgYDVR0PAQH/BAQDAgEGMCkGA1Ud
# EQQiMCCkHjAcMRowGAYDVQQDExFTeW1hbnRlY1BLSS0xLTU2ODAdBgNVHQ4EFgQU
# V4abVLi+pimK5PbC4hMYiYXN3LcwHwYDVR0jBBgwFoAUe1tFz6/Oy3r9MZIaarbz
# RutXSFAwDQYJKoZIhvcNAQELBQADggEBACQ79degNhPHQ/7wCYdo0ZgxbhLkPx4f
# lntrTB6HnovFbKOxDHtQktWBnLGPLCm37vmRBbmOQfEs9tBZLZjgueqAAUdAlbg9
# nQO9ebs1tq2cTCf2Z0UQycW8h05Ve9KHu93cMO/G1GzMmTVtHOBg081ojylZS4mW
# CEbJjvx1T8XcCcxOJ4tEzQe8rATgtTOlh5/03XMMkeoSgW/jdfAetZNsRBfVPpfJ
# vQcsVncfhd1G6L/eLIGUo/flt6fBN591ylV3TV42KcqF2EVBcld1wHlb+jQQBm1k
# IEK3OsgfHUZkAl/GR77wxDooVNr2Hk+aohlDpG9J+PxeQiAohItHIG4wggSjMIID
# i6ADAgECAhAOz/Q4yP6/NW4E2GqYGxpQMA0GCSqGSIb3DQEBBQUAMF4xCzAJBgNV
# BAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3JhdGlvbjEwMC4GA1UEAxMn
# U3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBDQSAtIEcyMB4XDTEyMTAx
# ODAwMDAwMFoXDTIwMTIyOTIzNTk1OVowYjELMAkGA1UEBhMCVVMxHTAbBgNVBAoT
# FFN5bWFudGVjIENvcnBvcmF0aW9uMTQwMgYDVQQDEytTeW1hbnRlYyBUaW1lIFN0
# YW1waW5nIFNlcnZpY2VzIFNpZ25lciAtIEc0MIIBIjANBgkqhkiG9w0BAQEFAAOC
# AQ8AMIIBCgKCAQEAomMLOUS4uyOnREm7Dv+h8GEKU5OwmNutLA9KxW7/hjxTVQ8V
# zgQ/K/2plpbZvmF5C1vJTIZ25eBDSyKV7sIrQ8Gf2Gi0jkBP7oU4uRHFI/JkWPAV
# Mm9OV6GuiKQC1yoezUvh3WPVF4kyW7BemVqonShQDhfultthO0VRHc8SVguSR/yr
# rvZmPUescHLnkudfzRC5xINklBm9JYDh6NIipdC6Anqhd5NbZcPuF3S8QYYq3AhM
# jJKMkS2ed0QfaNaodHfbDlsyi1aLM73ZY8hJnTrFxeozC9Lxoxv0i77Zs1eLO94E
# p3oisiSuLsdwxb5OgyYI+wu9qU+ZCOEQKHKqzQIDAQABo4IBVzCCAVMwDAYDVR0T
# AQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8EBAMCB4Aw
# cwYIKwYBBQUHAQEEZzBlMCoGCCsGAQUFBzABhh5odHRwOi8vdHMtb2NzcC53cy5z
# eW1hbnRlYy5jb20wNwYIKwYBBQUHMAKGK2h0dHA6Ly90cy1haWEud3Muc3ltYW50
# ZWMuY29tL3Rzcy1jYS1nMi5jZXIwPAYDVR0fBDUwMzAxoC+gLYYraHR0cDovL3Rz
# LWNybC53cy5zeW1hbnRlYy5jb20vdHNzLWNhLWcyLmNybDAoBgNVHREEITAfpB0w
# GzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMjAdBgNVHQ4EFgQURsZpow5KFB7V
# TNpSYxc/Xja8DeYwHwYDVR0jBBgwFoAUX5r1blzMzHSa1N197z/b7EyALt0wDQYJ
# KoZIhvcNAQEFBQADggEBAHg7tJEqAEzwj2IwN3ijhCcHbxiy3iXcoNSUA6qGTiWf
# mkADHN3O43nLIWgG2rYytG2/9CwmYzPkSWRtDebDZw73BaQ1bHyJFsbpst+y6d0g
# xnEPzZV03LZc3r03H0N45ni1zSgEIKOq8UvEiCmRDoDREfzdXHZuT14ORUZBbg2w
# 6jiasTraCXEQ/Bx5tIB7rGn0/Zy2DBYr8X9bCT2bW+IWyhOBbQAuOA2oKY8s4bL0
# WqkBrxWcLC9JG9siu8P+eJRRw4axgohd8D20UaF5Mysue7ncIAkTcetqGVvP6KUw
# VyyJST+5z3/Jvz4iaGNTmr1pdKzFHTx/kuDDvBzYBHUwggTrMIID06ADAgECAhAv
# iUKKNur0wAffSfaDzgbOMA0GCSqGSIb3DQEBCwUAMEwxCzAJBgNVBAYTAlVTMRUw
# EwYDVQQKEwx0aGF3dGUsIEluYy4xJjAkBgNVBAMTHXRoYXd0ZSBTSEEyNTYgQ29k
# ZSBTaWduaW5nIENBMB4XDTE1MDEyMDAwMDAwMFoXDTE3MDIwNDIzNTk1OVowcjEL
# MAkGA1UEBhMCQVUxETAPBgNVBAgMCFZpY3RvcmlhMRIwEAYDVQQHDAlNZWxib3Vy
# bmUxDzANBgNVBAoMBk5pbnRleDEaMBgGA1UECwwRQnVzaW5lc3MgU2VydmljZXMx
# DzANBgNVBAMMBk5pbnRleDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# ANoFNkVtFI7kVAkGkNVr77PxGEcd7JEfYGZy3cjY5rwWhqlcLFHeoNxK1uJ8nlXG
# 6QQbfx6P97R7uUkHDjYhfV5GBIHy/fXNEAbMxYWLvVMUZx3cdstEKeSHe8EstvBq
# dSlo1gDCH1RIPHbf8OZFXMvGBYHGoTDrgXviM501I61Xg2gx4AGJLDNYfmaSXTW0
# cpdyxs4eecF7h+Ey/TkE7oZhcZRqgwu+/mopldadrSHJCtnDcKyHnw+Y8+UXr9fN
# IoZwrdryHU77dy0aswloCbZWBEuXXJprzZjuFyKdsrHM6sBSTe7AeWCza7vlxMlw
# z4+GexPo3ZTw0z5VXpZYLLkCAwEAAaOCAaEwggGdMAkGA1UdEwQCMAAwHwYDVR0j
# BBgwFoAUV4abVLi+pimK5PbC4hMYiYXN3LcwHQYDVR0OBBYEFIJGZ/uH3FIA8tQD
# J6aQs/ZKSsWCMCsGA1UdHwQkMCIwIKAeoByGGmh0dHA6Ly90bC5zeW1jYi5jb20v
# dGwuY3JsMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzBzBgNV
# HSAEbDBqMGgGC2CGSAGG+EUBBzACMFkwJgYIKwYBBQUHAgEWGmh0dHBzOi8vd3d3
# LnRoYXd0ZS5jb20vY3BzMC8GCCsGAQUFBwICMCMMIWh0dHBzOi8vd3d3LnRoYXd0
# ZS5jb20vcmVwb3NpdG9yeTAdBgNVHQQEFjAUMA4wDAYKKwYBBAGCNwIBFgMCB4Aw
# VwYIKwYBBQUHAQEESzBJMB8GCCsGAQUFBzABhhNodHRwOi8vdGwuc3ltY2QuY29t
# MCYGCCsGAQUFBzAChhpodHRwOi8vdGwuc3ltY2IuY29tL3RsLmNydDARBglghkgB
# hvhCAQEEBAMCBBAwDQYJKoZIhvcNAQELBQADggEBAJJjCOjasSPkBwmC3680Ifjj
# jr9H2kzJ3rcroHoJLfV4K1HngUob8+zanVatHDjQRuGXkBKWgKzZODEhgM64mjFc
# PrxXlvwsJRO8G43bf4Ni/ztgim1Ze9Id/30kmsVQMYTe+/JB2J/y/hBsQ3e7LzdI
# YwK5cCe1D2r9ewTV44oKSGeAKuMrYXbVljxUaTM3yOn40Fy3zgok4WNNUuu2od5m
# Kklo8NUQbdktBbM3wSZFt4Mw9H0C7v/owi5Y3BI9VkJn2LoX8GB3d+wUE8XyUzPg
# QDOPoz5f+InjkO2McMsDt6AMnLQxxYEVkDlB1O0o9xR6dTn/BdFkfxaAKt8KssQx
# ggQQMIIEDAIBATBgMEwxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwx0aGF3dGUsIElu
# Yy4xJjAkBgNVBAMTHXRoYXd0ZSBTSEEyNTYgQ29kZSBTaWduaW5nIENBAhAviUKK
# Nur0wAffSfaDzgbOMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgACh
# AoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAM
# BgorBgEEAYI3AgEWMCMGCSqGSIb3DQEJBDEWBBQI9kpQQPAbLQRkPDfhT+yz9+a6
# ZjANBgkqhkiG9w0BAQEFAASCAQA8OFveXMeOlK0xbpWthI3nCAGRAP6ZXnMDOEEy
# AcaeQzwVSVmiBT32gkRb1yuLQ3Bku70qQ9TIdztY1LMZbhM6j1ZUHlh9eruM/qHK
# lqscVGiVTHyvG4ZRDVEZXd33fPjNZ7Q0fdIGOjUyQ7mcnkggbWlH/E0DbEGetgE6
# FhtcTdGJtFEccSdJSyP1inuIUEO9NSlhUWGTQpQQjXeO3NYKuUqbWxnstqGzi1gW
# 5PycZzxX96td2D0rS5nsrPXP920qRLIt1kuWuLguJB8kDItHGGC7KHSovUEyg4mx
# sASK60On7Ot8Hvka9t35+WzkyHfrPUAP1Mdn8R619QM3x0PeoYICCzCCAgcGCSqG
# SIb3DQEJBjGCAfgwggH0AgEBMHIwXjELMAkGA1UEBhMCVVMxHTAbBgNVBAoTFFN5
# bWFudGVjIENvcnBvcmF0aW9uMTAwLgYDVQQDEydTeW1hbnRlYyBUaW1lIFN0YW1w
# aW5nIFNlcnZpY2VzIENBIC0gRzICEA7P9DjI/r81bgTYapgbGlAwCQYFKw4DAhoF
# AKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE1
# MDMyNzA1Mjc0OFowIwYJKoZIhvcNAQkEMRYEFDozuUhd9TYNnu5WlwBMUC73YKNE
# MA0GCSqGSIb3DQEBAQUABIIBAIzKY4XhbqrdYyCwW2zvkzoY/cmp0cakkqlZyauB
# kLWLjQ0soy3G1glUVHZJA6eDLNSxAqlR4f3G0DwqLlZH7vCxiCXbf5/NF13vgK0r
# cYaesy07bpezcAibTslEE8f6tCzkLq1d5bwD1gpHmbcLBNt6qghYvGSh70EfMUNI
# iqCdWr2LREfes4HH+fIHq5FSRhVKhLlnMtZRlj6OuY0dMwr/XpiMBR23c6/EO+0Y
# q4JFud+XFsJMavhuv+QHzeW+1e1cWSpGyx8RXF2JcqToCZCNcbOZbOJOkPczTLSV
# 0jBsI2HewVdEowR3HR77UJfPT6y7xFlmN9Zh8u5YHhz+Psk=
# SIG # End signature block
