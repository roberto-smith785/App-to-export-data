$sourceScript = "C:\Users\rober\OneDrive\Pictures\powershell\app.ps1"
$outputExe = "C:\Users\rober\OneDrive\Pictures\powershell\app.exe"
$logo = "C:\Users\rober\OneDrive\Pictures\powershell\app.ico"
# Corrected PS2EXE command
Invoke-ps2exe $sourceScript $outputExe -iconFile $logo -noConsole
