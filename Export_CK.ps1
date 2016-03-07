############################################################################################################################  
# Powershell Script:    Export_MPs.ps1  
#  
# Author:  SCOM Admin
#  
# Inception:            29.01.2016  
# Last Modified:        29.01.2016 
#  
# Description:          This Script reads all the SCOM management pack Company Knowledge and exports it to a .html file format.  Its useful if you use updated and use SCOM Company Knowedge for alerting or ticketing purpose.
#  
# Version Updates:       
#			 
#  
# PowerShell Syntax:    ./Export_CK.ps1 <MP Name> <String Pattern to search in CK>  <Outfile Name.html>
#  
# PowerShell Examples:   .\Export_CK.ps1  "ALL"  "" "CK_All.html"    --> Extracts company knowledge for all products
#                        .\Export_CK.ps1  "Microsoft.SQLServer.2008.Monitoring"  "" "CK_All.html" --> Extracts company knowledge only for a specfic SCOM management pack'
#                        .\Export_CK.ps1  "Microsoft.SQLServer.2008.Monitoring" "Priority: High" "C:\Tools\Output.html" --> Extracts company knowledge only for SQL 2008 product and gets only CK containing Priority: High '
#                        .\Export_CK.ps1  "Microsoft.SQLServer.2008.Monitoring" "Priority: Low" "C:\Tools\Output.html" --> Extracts company knowledge only for SQL 2008 product and gets only CK containing Priority: Low '
#  
############################################################################################################################  
 
 

function fnHtmlToCleanHTML($pHtmlText, $Search_pattern)
{

  $pHtmlText = $pHtmlText -replace("&lt;","<");
  $pHtmlText = $pHtmlText -replace("&gt;",">");

  if (($pHtmlText.contains("<html><body></body></html>") -eq 0 ) -and ($pHtmlText))
  {
    $pDIV = "<div style='mso-element:para-border-div;border:none;border-bottom:solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .75pt;padding:0cm 0cm 1.0pt 0cm'></div>"
    $heading = "<p><B>Management Pack Name:</B>&nbsp;&nbsp;" + $MP.DisplayName + "&nbsp;&nbsp;<B>Enabled:</B>&nbsp;&nbsp;" +  $_.Enabled + "&nbsp;&nbsp;<B>Monitor\Rule Name:</B>&nbsp;&nbsp;" + $_.DisplayName  + "&nbsp;&nbsp;<B>Category:</B>&nbsp;&nbsp;" + $_.Category
    $Search_pattern
    $pHtmlText

    $pHtmlText = $pHtmlText + $pDIV

    $blnFound = $FALSE

    if ($pHtmlText.count -gt 1)
    {
        for ($i=0;$i -lt $pHtmlText.count;$i++)
        {
            if ($pHtmlText[$i].Contains($Search_pattern) -gt 0)
            {
               $blnFound = $TRUE
            }
        }
     }
     else
     {
            if ($pHtmlText.Contains($Search_pattern) -gt 0)
            {
               $blnFound = $TRUE
            }

     }


    if (($blnFound -eq $TRUE) -or ($Search_pattern  -eq ""))
    {
        Write-Host "Found $Search_pattern) in CK $pHtmlText"
        
        $heading | Out-File $sOutPutFile -Append
        $pHtmlText | Out-File $sOutPutFile -Append
    }
  }
}


function fnMamlToHTML($MAMLText)
{
 $HTMLText = "";
 $HTMLText = $MAMLText -replace('xmlns:maml="http://schemas.microsoft.com/maml/2004/10"');
 $HTMLText = $HTMLText -replace("maml:para","p");
 $HTMLText = $HTMLText -replace("maml:");
 $HTMLText = $HTMLText -replace("</section>");
 $HTMLText = $HTMLText -replace("<section>");
 $HTMLText = $HTMLText -replace("<section >");
 $HTMLText = $HTMLText -replace("<title>","<h2>");
 $HTMLText = $HTMLText -replace("</title>","</h2>");
 $HTMLText = $HTMLText -replace("<listitem>","<li>");
 $HTMLText = $HTMLText -replace("</listitem>","</li>");
 $HTMLText = $HTMLText -replace("<navigationLink>");
 $HTMLText = $HTMLText -replace("</navigationLink>");
 $HTMLText = "<html><body>" + $HTMLText + "</body></html>";
 $HTMLText;
}

function fnMamlToVariable($MAMLText)
{
 $HTMLText = "";
 $HTMLText = $MAMLText -replace('xmlns:maml="http://schemas.microsoft.com/maml/2004/10"');
 $HTMLText = $HTMLText -replace("maml:para","p");
 $HTMLText = $HTMLText -replace("maml:");
 $HTMLText = $HTMLText -replace("</section>");
 $HTMLText = $HTMLText -replace("<section>");
 $HTMLText = $HTMLText -replace("<section >");
 $HTMLText = $HTMLText -replace("<title>","");
 $HTMLText = $HTMLText -replace("</title>","");
 $HTMLText = $HTMLText -replace("<listitem>","");
 $HTMLText = $HTMLText -replace("</listitem>","");
 $HTMLText;
}


Function GetAllCK()
{
        if (test-path ( $sOutPutFile))
        {
           Remove-Item $sOutPutFile
        }

        $mg = New-Object Microsoft.EnterpriseManagement.ManagementGroup("xwnscommgmt02")
        $cul = "ENU"

        Get-SCOMManagementPack   | ForEach-Object `
        {
            $strMPName = $_.Name

            if (($Search_MP_Name -eq  $strMPName) -or ($Search_MP_Name -eq "ALL"))
            {
                Write-Host "Reading MP DisplayName "   $_.DisplayName  " ==> Mp Name "  $_.Name
            
                $MP = Get-SCOMManagementPack -Name $_.Name
                $i = 0

                $all_monitors = Get-SCOMMonitor -ManagementPack  $MP
                $all_monitors | ForEach-Object  `
                { 
                    $maml = $mg.GetMonitoringKnowledgeArticles($_.Id).HTMLcontent
    
                    if ($maml.Trim.length -gt 0)
                    {
                           fnHtmlToCleanHTML $maml $Search_pattern
                    }
                    $i = $i + 1
                }


                $i = 0
                $all_rules = Get-SCOMRule -ManagementPack  $MP
                $all_rules | ForEach-Object  `
                { 
                $maml = $mg.GetMonitoringKnowledgeArticles($_.Id).HTMLcontent

               
                if ($maml.Trim.length -gt 0)
                {
                    if ($Search_pattern.length -eq 0)
                    { 
                        fnHtmlToCleanHTML $maml $Search_pattern
                    }
                }
                $i = $i + 1

                }
            }
         }
}

Import-Module OperationsManager
$blnDebug = 0
$Search_MP_Name = ""
$Search_pattern = ""
$sOutPutFile = ""

if ($blnDebug -eq 1)
{ $Search_MP_Name= "ALL"
  $Search_pattern = ""
  $sOutPutFile = "test.html"
  GetAllCK
}

$mp = ""
if ($args.Count -eq 3)
{
    $Search_MP_Name = $args[0]
    $Search_pattern = $args[1]
    $sOutPutFile = $args[2]
    Write-Host "Reading CK and searching $Search_MP_Name containing $Search_pattern"
    GetAllCK
    if (test-path ( $sOutPutFile))
    {
        Write-Host "Output file created " $sOutPutFile 
    }
    else
    {
        Write-Host "Opps!!Output file not created or no records found with the given criteria."
    }
}
else
{ 
    Write-Host "Syntax Error!!"
    Write-Host "The syntax for getting the monitors is: Export_CK.ps1 <MP Name> <String Pattern to search in CK>  <Outfile Name.html>" 
    Write-Host '.\Export_CK.ps1 "ALL"  "" "CK_All.html"    --> Extracts company knowledge for all products'
    Write-Host '.\Export_CK.ps1 "Microsoft.SQLServer.2008.Monitoring"  "" "CK_All.html" --> Extracts company knowledge only for a specfic SCOM management pack'
    Write-Host '.\Export_CK.ps1  "Microsoft.SQLServer.2008.Monitoring" "Priority: High" "C:\Tools\Output.html" --> Extracts company knowledge only for SQL 2008 product and gets only CK containing Priority: High '
    Write-Host '.\Export_CK.ps1  "Microsoft.SQLServer.2008.Monitoring" "Priority: Low" "C:\Tools\Output.html" --> Extracts company knowledge only for SQL 2008 product and gets only CK containing Priority: Low '
    Write-Host 
}
