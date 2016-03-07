############################################################################################################################ 
# Powershell Script:    Export_MPs.ps1 
# 
# Author:  Rajesh Naik
# 
# Inception:            29.01.2016 
# Last Modified:        29.01.2016
# 
# Description:          This Script reads all the SCOM management pack and exports it to a .CSV file on a given folder.
# 
# Version Updates:      
#			
# 
# PowerShell Syntax:    ./Export_MS.ps1 "DRIVE_Name:\FOLDER_Path"
# 
# PowerShell Example:   ./Export_MS.ps1 "C:\Users\userid\Desktop\Scripts\Tools\ExportMPs"
# 
############################################################################################################################ 

$all_mps = get-SCOMmanagementpack
 
foreach($mp in $all_mps)
 
{
 
Export-SCOMManagementPack -managementpack $mp -path "C:\Users\userid\Desktop\Scripts\Tools\ExportMPs"
 
} 
