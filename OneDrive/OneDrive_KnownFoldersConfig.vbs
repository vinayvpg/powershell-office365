' Set up all strings**
' This script is Murphy Oil Corporation specific
' Created by Brent Jackson - Sparkhound 7 July 2017
' Modified to support manual folder redirection
Option Explicit
Public Const HKEY_CURRENT_USER = &H80000001
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Public ObjFSO, objFile, logFile, objShell,regKeyPath, regUserShellPath, oReg, objUserEnv
Public strEmail,strOneDriveSetup, objADSysInfo, arrValueNames, arrValueTypes, objUser
Public regDocs, regFav, regDesktop, regPics, regVideos, upper, i, regOneDrive
Public actualDocs, actualFav, actualDesktop, actualPics, actualVideos, strRoboCALDocs
Public strUserName, counter, strOneDrive, strUserProfile, strValueName,strCALDocs
Public Const strComputer = "."
Public strDate, strTempFolder, boolFolderExists, boolCopyComplete, readme
Dim strRoboFav, strRoboDocs, strRoboPics, strRoboDesk, strRoboVideos
WScript.Sleep 10000
' Set up the log file with dates
strDate=strDate &"\" & Year(Now) & "." & Month(Now) & "." & Day(Now) & "."
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")	
Set objUserEnv = objShell.Environment("USER")
strTempFolder = objShell.ExpandEnvironmentStrings("%TEMP%")
strUserProfile=objShell.ExpandEnvironmentStrings("%USERPROFILE%")
strUserName=objShell.ExpandEnvironmentStrings("%USERNAME%")
strOneDrive=strUserProfile & "\OneDrive - Murphy Oil"
strCALDocs="\\calcifs01\users\" & strUsername

actualDocs= strOneDrive & "\Documents"
actualFav= strOneDrive & "\Favorites"
actualPics= strOneDrive & "\Pictures"
actualDesktop= strOneDrive & "\Desktop"
actualVideos= strOneDrive & "\Videos"

'set up log file for appending
logFile=strTempFolder & strDate & "OneDriveConfiguration.log"
Set objFile = objFSO.OpenTextFile(logFile,ForAppending, True)
objFile.Write "****************** OneDrive Configuration *************************" & vbCrLf

'set up registry access
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
regKeyPath = "SOFTWARE\Microsoft\OneDrive\Accounts\Business1\ScopeIdToMountPointPathCache"
regUserShellPath="SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"

'write retrieved data to log
objFile.Write "User Name 		: " & strUserName & vbCrLf
objFile.Write "User Profile		: " & strUserProfile & vbCrLf
objFile.Write "OneDrive Profile	: " & strOneDrive & vbCrLf

'Check existence of Murphy Oil registry key if not configure OneDrive


On Error Resume Next
	oReg.EnumValues HKEY_CURRENT_USER, regKeyPath, arrValueNames, arrValueTypes
	Upper = UBound(arrValueNames)
On Error Goto 0

If Upper < 0 Or IsNull(Upper) Or IsEmpty(Upper) Then
	strValueName = "Does not exist"
Else
	strValueName = arrValueNames(0)
End If

objFile.Write "OneDrive Registry Name  : " & strValueName & vbCrlf

'if OneDrive is configured write to log and exit, otherwise configure OneDrive on client
If strValueName <> "Does not exist" Then
	regOneDrive = GetRegValue(regKeyPath, strValueName)
	objFile.Write "OneDrive Reg Value	: " & regOneDrive & vbCrLf
 	If strUserName = "Brent_Jackson" Then
 		MsgBox "OneDrive Key: "  & strValueName & vbCrLf & " i = " & i & vbCrLf & "OneDrive setting = " & regOneDrive, vbOkOnly,"OneDrive Client Configuration"
	End If

Else

	objFile.Write "Registry key for Murphy Oil NOT found - executing OneDrive configuration" & vbCrLf
	Set objADSysInfo = CreateObject("ADSystemInfo")
	DO
		Set objUser = GetObject("LDAP://" & objADSysInfo.UserName)
		strUserName = objUser.samaccountname
		strEmail = objUser.mail
	Loop Until InStr (strEmail, "murphyoilcorp.com") >0
	objFile.Write "Samaccountname	: "& strUserName & vbCrLf
	objFile.Write "Email Address	: "& strEmail & vbCrLf
	If strUserName = "Brent_Jackson" Then
		MsgBox "UserName: " & strUserName & vbCrLf & "Email: " & strEmail, vbokonly, "Found the following user"
	End If
	strOneDriveSetup="odopen://sync?useremail="+strEmail
	On Error Resume Next
	objFile.Write "OneDrive Setup executed for " & strEmail  & vbCrLf	
	objShell.Run strOneDriveSetup
	
End If

On Error Goto 0

'Check and validate favorites
Do
	'MsgBox "User Shell Path = " & regUserShellPath & vbCrLf & "UserProfile = " & strUserProfile, vbOKOnly,"Get Favorites location"
	regFav= GetRegValue(regUserShellPath, "Favorites")

	boolFolderExists=objFSO.FolderExists(actualFav)
	If strUserName = "Brent_Jackson" Then
		MsgBox "Folder exists = " & boolFolderExists & vbCrLf & "regFav = " & regFav & vbCrLf & "AcualFav = " & actualFav , vbOkOnly,"Favorites Validation"
	End If
Loop Until boolFolderExists And regFav = actualFav

'Move Favorites

If objFSO.FolderExists(strUserProfile & "\Favorites") Then
	objFile.Write "Moving Favorites" & vbCrLf
	strRoboFav = "C:\Windows\System32\Robocopy.exe """ & strUserProfile & "\Favorites"" """ & strOneDrive & "\Favorites"" /is /s /e /z /MOVE /r:1 /w:1 /TEE /LOG+:" & strTempFolder & "\Move-Favorites.log"
	objFile.Write strRoboFav & vbCrLf
	objShell.Run strRoboFav
End If
	
'check and validate Desktop
Do
	regDesktop=GetRegValue(regUserShellPath, "Desktop")
	boolFolderExists=objFSO.FolderExists(actualDesktop)
	If strUserName = "Brent_Jackson" Then
		MsgBox "Folder exists = " & boolFolderExists & vbCrLf & "regDesktop= " & regDesktop & vbCrLf & "AcualDesktop = " & actualDesktop, vbOkOnly,"Desktop Validation"
	End If
Loop Until boolFolderExists And regDesktop = actualDesktop

'move Desktop
If objFSO.FolderExists(strUserProfile & "\Desktop") Then
	objFile.Write "Moving Desktop" & vbCrLf
	strRoboDesk = "C:\Windows\System32\Robocopy.exe """ & strUserProfile & "\Desktop"" """ & strOneDrive & "\Desktop"" /is /s /e /z /MOVE /r:1 /w:1 /xf readme.murphy.txt /TEE /LOG+:" & strTempFolder & "\Move-Desktop.log"
	objFile.Write strRoboDesk & vbCrLf
	objShell.Run strRoboDesk
	Set readme = objFSO.CreateTextFile (strUserProfile & "\Desktop\readme.murphy.txt", True)
	readme.WriteLine "Per the new Murphy Oil OneDrive Policy your Desktop icons have now been moved to a new location."
	readme.WriteLine "Please log off and log back on and all will be restored."
	readme.WriteLine "Thank you."
	readme.Close
End If

'check and validate Documents
Do
	regDocs=GetRegValue(regUserShellPath, "Personal")
	boolFolderExists=objFSO.FolderExists(actualDocs)
	If strUserName = "Brent_Jackson" Then
		MsgBox "Folder exists = " & boolFolderExists & vbCrLf & "regDocs= " & regDocs & vbCrLf & "AcualDocs = " & actualDocs, vbOkOnly,"Documents Validation"
	End If
Loop Until boolFolderExists And regDocs = actualDocs

'move documents
	objFile.Write "Moving Documents" & vbCrLf
	strRoboDocs = "C:\Windows\System32\Robocopy.exe """ & strUserProfile & "\Documents"" """ & strOneDrive & "\Documents"" /is /s /e /z /MOVE /r:1 /w:1 /XD PGP SAP music videos pictures 'my music' 'my videos' 'my pictures' /xf *.pst *.tmp /TEE /LOG+:" & strTempFolder & "\Move-Documents.log"
	objFile.Write strRoboDocs & vbCrLf
	objShell.Run strRoboDocs


'check and validate Pictures
Do
	regPics=GetRegValue(regUserShellPath, "My Pictures")
	boolFolderExists=objFSO.FolderExists(actualPics)
	If strUserName = "Brent_Jackson" Then
		MsgBox "Folder exists = " & boolFolderExists & vbCrLf & "regPics= " & regPics & vbCrLf & "actualPics = " & actualPics, vbOkOnly,"Pictures Validation"
	End If
Loop Until boolFolderExists And regPics = actualPics


'move pictures
If objFSO.FolderExists(strUserProfile & "\Pictures") Then
	objFile.Write "Moving Pictures" & vbCrLf
	strRoboPics = "C:\Windows\System32\Robocopy.exe """ & strUserProfile & "\Pictures"" """ & strOneDrive & "\Pictures"" /is /s /e /MOVE /z /r:1 /w:1 /TEE /LOG+:" & strTempFolder & "\Move-Pictures.log"
	objFile.Write strRoboPics & vbCrLf
	objShell.Run strRoboPics
End If
	
'check and validate Videos
Do
	regVideos=GetRegValue(regUserShellPath, "My Video")
	boolFolderExists=objFSO.FolderExists(actualVideos)
	If strUserName = "Brent_Jackson" Then
		MsgBox "Folder exists = " & boolFolderExists & vbCrLf & "regVideo= " & regVideos & vbCrLf & "actualVideos = " & actualVideos, vbOkOnly,"Videos Validation"
	End If
Loop Until boolFolderExists And regVideos = actualVideos

'move videos
If objFSO.FolderExists(strUserProfile & "\Videos") Then
	objFile.Write "Moving Videos" & vbCrLf
	strRoboVideos = "C:\Windows\System32\Robocopy.exe """ & strUserProfile & "\Videos"" """ & strOneDrive & "\Videos"" /is /s /e /MOVE /z /r:1 /w:1 /TEE /LOG+:" & strTempFolder & "\Move-Videos.log"
	objFile.Write strRoboVideos & vbCrLf
	objShell.Run strRoboVideos
End If

objFile.Write "----------------- Ending configuration ----------------"
objFile.Close
Set oReg = Nothing
Set objShell = Nothing
Set objFSO = Nothing
Set objADSysInfo = Nothing
'Function to get registry keys
Function GetRegValue(regPath, regKeyName)
	Dim regLocation
	oReg.GetExpandedStringValue HKEY_CURRENT_USER,regPath, regKeyName, regLocation
	'regLocation= objShell.RegRead ("HKEY_CURRENT_USER\" & regUserShellPath & "\" & regKeyName)
	regLocation = Replace(regLocation, "%USERPROFILE%", strUserProfile )

	If strUserName = "Brent_Jackson" Then
		MsgBox "HKCU = " & HKEY_CURRENT_USER & vbcrlf & "RegPath = " & regPath & vbCrlf & "regKeyName = " & regKeyName & vbCrLf & "GetRegValue = " & regLocation, vbokonly, "Inside Function call"
	End If
	GetRegValue = regLocation
End Function
