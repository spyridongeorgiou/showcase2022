# Create a Shortcut with Windows PowerShell
$SourceFileLocation = "C:\EDV\Lesehilfe\Lesehilfe.exe"
$ShortcutLocation = "C:\Users\Public\Desktop\Lesehilfe.lnk"
#New-Object : Creates an instance of a Microsoft .NET Framework or COM object.
#-ComObject WScript.Shell: This creates an instance of the COM object that represents the WScript.Shell for invoke CreateShortCut
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutLocation)
$Shortcut.TargetPath = $SourceFileLocation
#Save the Shortcut to the TargetPath
$Shortcut.Save()