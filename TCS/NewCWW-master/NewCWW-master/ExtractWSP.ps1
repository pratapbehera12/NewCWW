## SharePoint DLL 
Add-PSSnapin "Microsoft.SharePoint.Powershell" -ErrorAction SilentlyContinue 
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
Start-SPAssignment –Global

$dirName = "c:\Solutions"
foreach ($solution in Get-SPSolution)
{
    $id = $Solution.SolutionID
    $title = $Solution.Name
    $filename = $Solution.SolutionFile.Name
    $solution.SolutionFile.SaveAs("$dirName\$filename")
}


#Read more: http://www.sharepointdiary.com/2011/10/extract-download-wsp-files-from-installed-solutions.html#ixzz5QeevZ3lo