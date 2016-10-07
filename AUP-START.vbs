' AUP by Ian Mullings
' Please see https://github.com/imullings/AUP

On Error resume next    
Set ws = CreateObject("wscript.Shell")
strHomeshare=ws.ExpandEnvironmentStrings("%HOMESHARE%")
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("AUP-settings.inf")
Do Until f.AtEndOfStream
	MyLine = f.ReadLine
 

 	'#######Read number of nrandomlogons
	if InStr(MyLine, "nrandomlogons") <> 0 and nrandomlogonsFound <> 1 then
    SettingsFind1 = InStr(MyLine, """")  
    SettingsFind2 = Mid(MyLine,SettingsFind1 + 1)
    SettingsFind3 = InStr(SettingsFind2, """")
	
    nrandomlogons = Mid(MyLine,SettingsFind1 + 1,SettingsFind3 - 1)
    nrandomlogonsFound = 1
   

	end if 
	'########
 	'#######Read path nPathToAcceptFile
	if InStr(MyLine, "nPathToAcceptFile") <> 0 and nPathToAcceptFileFound <> 1 then
    SettingsFind1 = InStr(MyLine, """")  
    SettingsFind2 = Mid(MyLine,SettingsFind1 + 1)
    SettingsFind3 = InStr(SettingsFind2, """")
	
    nPathToAcceptFile = Mid(MyLine,SettingsFind1 + 1,SettingsFind3 - 1)

    nPathToAcceptFileFound = 1
   
	
	end if 
	'########
	 	'#######Read nAUPTextVersion
	if InStr(MyLine, "nAUPTextVersion") <> 0 and nAUPTextVersionFound <> 1 then
    SettingsFind1 = InStr(MyLine, """")  
    SettingsFind2 = Mid(MyLine,SettingsFind1 + 1)
    SettingsFind3 = InStr(SettingsFind2, """")
	
    nAUPTextVersion = Mid(MyLine,SettingsFind1 + 1,SettingsFind3 - 1)
    nAUPTextVersionFound = 1
   
	
	end if 
	'########
Loop
f.Close

	If nAUPTextVersion = "" then
	nAUPTextVersion = "1"
	end if
	If nPathToAcceptFile = "" then
	nPathToAcceptFile = strHomeshare & "\accept" & nAUPTextVersion & ".htm"
	else
	nPathToAcceptFile = nPathToAcceptFile & "\accept" & nAUPTextVersion & ".htm"
	end if
	If nrandomlogons = "" then
	nrandomlogons = "20"
	end if	



' Check if file exists in homeshare
 Set objFSO = CreateObject("Scripting.FileSystemObject")
 If objFSO.FileExists(nPathToAcceptFile) Then
	'if it does only popup at random time
	Randomize
	randomnumber = Int(nrandomlogons * Rnd()) 
	if randomnumber = 1 then
	Show = 1
	end if
 Else
     Show = 1
 End If
 If show = 1 or nrandomlogons = 1 or  nrandomlogons = 0 then
       Set WshShell = CreateObject("WScript.Shell") 
        Return = WshShell.Run("AUP-HTA.hta",1,True) 
End If