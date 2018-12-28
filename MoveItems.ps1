##############################################
##                                          ##
##               Written by                 ##
##                                          ##
##               Ben Brandes                ##
##                                          ##
##############################################


#Adding outlook assembly to powershell
Add-Type -AssemblyName microsoft.office.interop.outlook

#create instance of outlook
$outlook = New-Object -ComObject outlook.application
$namespace = $Outlook.GetNameSpace("MAPI")

Write-Host "**************************************************************************"
Write-Host "**************************************************************************"
Write-Host "*********                                                        *********"
Write-Host "*********                  Outlook Items Mover                   *********"
Write-Host "*********                                                        *********"
Write-Host "**************************************************************************"
Write-Host "**************************************************************************"
Write-Host


###################    Source Folder Selection   ###################
Write-Host "#####   You're now going to select source folder   #####"
Write-Host "These are the main folders:"
For($i = 1; $i -lt ($namespace.Folders.Count + 1); $i++){
    Write-Host $i". " $namespace.Folders[$i].Name
}
$select = Read-Host "Choose Main Folder(by number)"

#save selected folder
Write-Host
$MainSelected = $namespace.Folders[[int]$select].Folders

Write-Host "These are the sub folders:"
For($i = 1; $i -lt ($MainSelected.Count + 1); $i++){
    Write-Host $i". " $MainSelected[$i].Name
}
$select = Read-Host "Choose sub Folder(by number)"

#Choose the Source folder
$SourceFolder = $MainSelected[[int]$select]

Write-Host
###################    Destination Folder Selection   ###################
Write-Host "#####   You're now going to select destination folder   #####"
Write-Host "These are the main folders:"
For($i = 1; $i -lt ($namespace.Folders.Count + 1); $i++){
    Write-Host $i". " $namespace.Folders[$i].Name
}
$select = Read-Host "Choose Main Folder(by number)"

#save selected folder
Write-Host
$MainSelected = $namespace.Folders[[int]$select].Folders

Write-Host "These are the sub folders:"
For($i = 1; $i -lt ($MainSelected.Count + 1); $i++){
    Write-Host $i". " $MainSelected[$i].Name
}
$select = Read-Host "Choose sub Folder(by number)"

#Choose the Destination folder
$DestFolder = $MainSelected[[int]$select]

Write-Host
Write-Host "Moving from " $SourceFolder.FolderPath
Write-Host "To " $DestFolder.FolderPath

$mails = $SourceFolder.Items

$i = 0;
$count = $mails.Count

$progress
For($i=$count; $i -gt 0; $i--){
    $mails[$i].Move($DestFolder) | Out-Null
    $progress++
    Write-Progress -activity "Moving items..." -status "Moved: $progress of $($count)" -percentComplete (($progress / $count)  * 100)
}
