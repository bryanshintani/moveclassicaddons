### Extracts .zip files in .\ to C:\Program Files (x86)\World of Warcraft\_classic_\Interface\AddOns
### and deletes the .zip files in .\
### Example: Download addons from curseforge.com to C:\Users\username\Downloads\WowAddons 
###          Place moveaddon.ps1 into same directory and run with PowerShell


#Check if WoWAddons folder exists, if not, create it in \Downloads\

$downloadsPath = "C:\Users\$env:username\Downloads\WoWAddons"
Write-Host "Checking for '$downloadsPath' directory"
if (-not (Test-Path -LiteralPath "$downloadsPath")) {
    try {
    New-Item -Path "C:\$env:username\Downloads" -Name "ClassicWowAddons" -ItemType "directory"
    Write-Host "Successfully created '$downloadsPath' directory"
    }
    catch {
        Write-Error -Message "Unable to create directory. Error was $_" -ErrorAction Stop
    }
} 
else {
    Write-Host "Directory already exists`n"
}


#Check if \WoWAddons\moveclassicaddons.ps1 exists, if not, move it

$checkFile = "C:\Users\$env:username\Downloads\WoWAddons\moveclassicaddons.ps1"
if (-not (Test-Path $checkFile)) {
    try {
        Move-Item -Path .\moveclassicaddons.ps1 -Destination $downloadsPath
        Write-Host "Successfully moved 'moveclassicaddons.ps1' to $downloadsPath"
    }
    catch {
        Write-Error -Message "Unable to move file. Error was $_" -ErrorAction Stop
    }
}
else {
    Write-Host "'$checkFile' already exists"
}


#Expand .zip's, move to \Interface\AddOns', remove zip's

Get-ChildItem -Name | ForEach-Object {
    if (([IO.FileInfo]$_).Extension -like '.zip') {
        Expand-Archive -Path $_ -DestinationPath 'C:\Program Files (x86)\World of Warcraft\_classic_\Interface\AddOns'
        Remove-Item -Path $_
        Write-Host "Successfully moved " ([IO.FileInfo]$_).Name
    }
}

Read-Host "`nPress any key to exit"