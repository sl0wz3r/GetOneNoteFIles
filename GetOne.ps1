# Set variables for local folder to move files to and zip file name
$LocalFolder = "C:\Temp"
$ZipFileName = "OneNoteFiles.zip"

# Set variables for OneNote storage folders and user's Downloads folder
$OneNote2016Folder = "$env:LOCALAPPDATA\Microsoft\OneNote\16.0"
$OneNote2013Folder = "$env:LOCALAPPDATA\Microsoft\OneNote\15.0"
$OneNote2010Folder = "$env:APPDATA\Microsoft\OneNote\14.0"
$DownloadsFolder = "$env:USERPROFILE\Downloads"

# Search for all .one files in the OneNote storage folders and user's Downloads folder
$Files = Get-ChildItem -Path $OneNote2016Folder,$OneNote2013Folder,$OneNote2010Folder,$DownloadsFolder -Recurse -Filter *.one

# Move files to a local folder before deleting them
Move-Item -Path $Files.FullName -Destination $LocalFolder

# Zip files into a single zip file
$ZipFilePath = Join-Path $LocalFolder $ZipFileName
Compress-Archive -Path $LocalFolder\* -DestinationPath $ZipFilePath

# Delete the original files
Remove-Item $Files.FullName
