' SuperNova Shortcut Maker by nosamu

' Get desktop folder
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set wshShell = CreateObject("WScript.Shell")
desktopDir = wshShell.ExpandEnvironmentStrings("%HOMEDRIVE%%HOMEPATH%") & "\Desktop"

' Get the command-line arguments of the currently running Supernova Player
Set objFSO = CreateObject("Scripting.FileSystemObject")
tempFilePath = scriptDir & "SupernovaShortcutMakerTemp.txt"
WshShell.Run "CMD /C WMIC path win32_process get Commandline | findstr snlauncher.exe > """ + tempFilePath + """", 0, True
processesStr = objFSO.OpenTextFile(tempFilePath, 1, False).ReadLine
If InStr(processesStr, "supernova://") = 0 Then
	msgBox "SuperNova Player does not appear to be running. Please try again."
	WScript.Quit
End If

' Get the SuperNova protocol URL from the arguments
Set re = New RegExp
re.Pattern = "(.*)(supernova://.*)"
re.Global  = False
re.IgnoreCase = True
supernovaUrl = re.Replace(processesStr, "$2")
' Trim whitespace
re.pattern = "^\s+|\s+$"
supernovaUrl = re.Replace(supernovaUrl, "")

urlShortcutName = ""
Do While Not IsKosherFilename(urlShortcutName) 
	urlShortcutName = InputBox("Enter a name for your shortcut.", "Create SuperNova Shortcut")
	If urlShortcutName = "" Then
		WScript.Quit
	ElseIf Not IsKosherFilename(urlShortcutName) Then
		msgBox "A shortcut cannot contain any of the following characters: \/:*?""<>|"
    End If
Loop

urlShortcutFile = desktopDir & "\" & urlShortcutName & ".url"
Set objFile = objFSO.CreateTextFile(urlShortcutFile, True)
objFile.Write "[InternetShortcut]" & vbCrLf
objFile.Write "URL=" & supernovaUrl
objFile.Close
msgBox "Shortcut saved!"

Function IsKosherFilename(fileName)
    If fileName = "" Then
        IsKosherFilename = false
		Exit Function
    End If

	Dim regEx
	Set regEx = New RegExp
	regEx.Pattern = "[\\\/:*?""<>|]"
	IsKosherFilename = NOT regEx.Test(fileName)
	Set RegEx = Nothing
End Function