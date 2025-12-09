' SamacSys Plugin for Altium Designer

Sub SamacSysPlugin
'    On Error Resume Next
    Dim InstalledDir, WshShell, fs
    Set WshShell = CreateObject("WScript.Shell")
    Set fs = CreateObject("Scripting.FileSystemObject")
    set InstallFolder = fs.GetFolder(WshShell.ExpandEnvironmentStrings(WshShell.SpecialFolders("MyDocuments")))
    InstalledDir =  InstallFolder.ShortPath & "\SamacSys\"
    WshShell.Run ("mshta.exe " + InstalledDir + "SamacSys.hta" & "")
End Sub
