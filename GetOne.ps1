# Set variables for file share path and local folder to move files to
$FileSharePath = "\\SERVER\SHARE\PATH"
$LocalFolder = "C:\Temp"

# Set variables for OneNote storage folders and user's Downloads folder
$OneNote2016Folder = "$env:LOCALAPPDATA\Microsoft\OneNote\16.0"
$OneNote2013Folder = "$env:LOCALAPPDATA\Microsoft\OneNote\15.0"
$OneNote2010Folder = "$env:APPDATA\Microsoft\OneNote\14.0"
$DownloadsFolder = "$env:USERPROFILE\Downloads"

# Search for all .one files in the OneNote storage folders and user's Downloads folder
$Files = Get-ChildItem -Path $OneNote2016Folder,$OneNote2013Folder,$OneNote2010Folder,$DownloadsFolder -Recurse -Filter *.one

# Loop through each file and copy it to the file share
foreach ($File in $Files) {
    $DestinationPath = Join-Path $FileSharePath $File.Name
    Copy-Item $File.FullName $DestinationPath
}

# Move files to a local folder before deleting them
Move-Item -Path $Files.FullName -Destination $LocalFolder

# Delete the original files
Remove-Item $Files.FullName
