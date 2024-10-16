$sourceScript = "C:\Github repos\App-to-export-data\app.ps1"
$outputExe = "C:\Github repos\App-to-export-data\app.exe"
$logo = "C:\Github repos\App-to-export-data\app.ico"
# Corrected PS2EXE command
Invoke-ps2exe $sourceScript $outputExe -iconFile $logo -noConsole
