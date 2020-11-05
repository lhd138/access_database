strDbPathName = CreateObject("Scripting.FileSystemObject").GetFolder(".").Path & "\Main.mdb"
CreateObject("WScript.Shell").Run "MSACCESS """ & strDbPathName & """ /decompile"