
#GAM CREATE FOLDER!
#Parent ID folder = 'Test GAM' on eecollabtest
## Automated M&A Discovery (Google)
$workingdir = "C:\WMScripts"
$gamcmd = "C:\WMScripts\gam.exe" 
$projectname = "Reginald"
$MAFolderID = "19zH0RuUi0xjWcy2SrXuDbUVSt8x9AL-p"
$adminuser = "adminuser@domain.com

#make sure Projectname First character is Upper case (Aesthetically pleasing)
$projectname = $projectname.Substring(0,1).toupper()+$projectname.Substring(1).tolower()

##------------ Top section is for JSON files needed for creation and updating------------------
# Create JSON for creating Sheet with first worksheet name being Users
$createsheet = @"
{"properties":{"title":"Project $projectname - Gsuite Discovery"},"sheets":[{"properties":{"sheetId":0,"title":"Users","index":0,"sheetType":"GRID"}}]}
"@

# Create JSON for Adding "Shared Drives" worksheet
$addsheet = @"
{"requests":[{"addSheet":{"properties":{"title":"Shared Drives","sheetType":"GRID"}}}]}
"@
###--- End JSON file ---
#Single Line JSON files


# Select Project source instance
Invoke-Expression "$gamcmd select $projectname save"
# Collect Users and Shared Drives from Source
Invoke-Expression "$gamcmd redirect csv $workingdir\$projectname-users.csv print users fields primaryemail,fullname,suspended"
Invoke-Expression "$gamcmd redirect csv $workingdir\$projectname-drives.csv print teamdrives asadmin fields name,id"

# Add Headers to CSV of both Drives and User Sheets
(Import-Csv .\$project-users.csv -Header "Primary Email", "Full Name", Suspended,Name,"Email Address","Dest Directory" | select -Skip 1) | Export-Csv .\$project-users.csv -NoTypeInformation
(Import-Csv .\$project-drives.csv -Header "Primary Email", "Full Name", Suspended,Name,"Email Address","Dest Directory" | select -Skip 1) | Export-Csv .\$project-drives.csv -NoTypeInformation

#Switch Back to Meta Production 
C:\WMScripts\gam.exe select fbprod save
# Create Project Folder and Folder Structure
if (!(((c:\wmscripts\gam.exe user thessellund@meta.com print filelist select teamdrive "EECOLLAB Test" corpora onlyshareddrives query "name = 'project $projectname'" fields id, name excludetrashed)|convertfrom-csv).id)) {
$projectfolder = c:\wmscripts\gam.exe user thessellund@meta.com create drivefile drivefilename "Project $projectname" mimetype gfolder parentid $MAFolderID
$projectfolderid = [regex]::Match($projectfolder,'(?<=\().+?(?=\))').value
foreach($item in "Discovery","Execution","Communications","Close"){
$folder = c:\wmscripts\gam.exe user thessellund@meta.com create drivefile drivefilename "$item" mimetype gfolder parentid $projectfolderid
if ($item -eq "Discovery") {
$discoveryfolderid = [regex]::Match($folder,'(?<=\().+?(?=\))').value
}
    
}
} Else { write-host "Folder Exists"}


#create Sheet for Discovery
$createsheet | Out-File $workingdir\sheet.json -Encoding utf8
$sheet = c:\wmscripts\gam.exe user $adminuser create sheet teamdriveparentid $discoveryfolderid json file $workingdir\sheet.json
#get sheetid for updating Sheet
$sheetid = [regex]::Match($sheet.get(1), '(?<=\: ).*$').value
# Adding "Shared Drives" worksheet
$addsheet | out-file $workingdir\sheet.json -Encoding utf8
C:\WMScripts\gam.exe user $adminuser update sheet $sheetid  json file $workingdir\sheet.json
#Populate Users Sheet
C:\WMScripts\gam.exe user $adminuser update drivefile id $sheetid retainname localfile $workingdir\$projectname-users.csv gsheet Users
#Populate Drives Sheet
C:\WMScripts\gam.exe user $adminuser update drivefile id $sheetid retainname localfile $workingdir\$projectname-drives.csv gsheet "Shared Drives"
