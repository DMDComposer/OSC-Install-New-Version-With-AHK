; v1.0.6
#Requires Autohotkey v1.1.33+
#SingleInstance, Force ; Limit one running version of this script
SetBatchlines -1 ; run at maximum CPU utilization
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
#Include %A_ScriptDir%\Includes\Jxon.ahk
#Include %A_ScriptDir%\Includes\m.ahk
#Include %A_ScriptDir%\Includes\Notify.ahk
; --------------------------------------

if (!A_IsAdmin){ ;http://ahkscript.org/docs/Variables.htm#IsAdmin
	MsgBox,, NOT ADMIN, You must run this script as an administrator, restarting app...
	Run *RunAs "%A_ScriptFullPath%" ; Requires v1.0.92.01+
	ExitApp
}

IniRead, oscPath, % A_ScriptDir "\settings.ini", settings, oscPath
global userOSCPath := checkUserOSCPath(oscPath)

currVersion := FileGetInfo(userOSCPath "\open-stage-control.exe").FileVersion

Endpoint := "https://api.github.com/repos/jean-emmanuel/open-stage-control/releases/latest"
req := ComObjCreate("WinHttp.WinHttpRequest.5.1")
req.Open("GET", Endpoint, true)
req.Send()
req.WaitForResponse()

if (req.status != 200) {
	MsgBox, % "Error finding latestVersion: " req.status
	return
}

; res := ConvertResponseBody(req)
res := req.ResponseBody
assets := JXON_Load(BinArr_ToString(res))
changelog := assets.body
latestVersion := SubStr(assets.tag_name, 2)

if (currVersion == latestVersion) {
	m("You are running the latest version of Open Stage Control")
	ExitApp
}

Loop % assets.assets.Length() {
	if (assets.assets[A_Index].name ~= "i)node.zip") {
		latestVersion := assets.assets[A_Index].name
		latestVersionURL := assets.assets[A_Index].browser_download_url
		Break
	}
}
if (!latestVersion) {
	m("couldn't find OSC latest version for Windows")
	ExitApp
}

MSGBox, 4, OSC Installing New Version , % latestVersion ": was found, would you like to install it?`nChanges Included:`n`n" changelog
IfMsgBox, No
Exit

Notify().AddWindow(latestVersion ": is installing now", {Title:"found latest version", Font:"Sans Serif", TitleFont:"Sans Serif", Icon:userOSCPath "\open-stage-control.exe", Animate:"Right, Slide", ShowDelay:100, IconSize:64, TitleSize:14, Size:20, Radius:26, Time:2500, Background:"0x2C323A", Color:"0xD8DFE9", TitleColor:"0xD8DFE9"})
; example URL: https://github.com/jean-emmanuel/open-stage-control/releases/download/v1.16.2/open-stage-control-1.16.2-win32-x64.zip
; UrlDownloadToFile, % "https://github.com/jean-emmanuel/open-stage-control/releases/download/v1.16.2/" latestVersion, % A_ScriptDir "/" latestVersion
UrlDownloadToFile, % latestVersionURL, % A_ScriptDir "/" latestVersion

if WinExist("Open Stage Control")												; Close OSC before installing new version
	WinClose, Open Stage Control

OSC_Zip := A_ScriptDir "\" latestVersion
OSC_Folder := userOSCPath "\resources\app"
7zip := "C:\Program Files\7-Zip"
SplitPath, OSC_Zip, vName, vDir, vEXT, vNNE, vDrive
if (vEXT != "zip"){ 														; If the OSC Zip isn't on clipboard, then exit app
	m("OSC Zip not on clipboard")
	ExitApp
}
FileDelete, % OSC_Folder "\*.*"
Loop, Files, % OSC_Folder "\*.*", D
{
	Path := OSC_Folder "\" A_LoopFileName
	if (Path != OSC_Folder "\_Archive") {
		FileRemoveDir, % Path, 1
	}
}

FileMove, % OSC_Zip, % OSC_Folder, 1
OSC_Zip_New := OSC_Folder . "\" . vName
While(!FileExist(OSC_Zip_New))
	Sleep, 10

DetectHiddenWindows, On
; Run, "%7zip%\7z.exe" x "%OSC_Zip_New%" -o"%OSC_Folder%"\ -y, % 7zip, Hide, PID
Unzip(OSC_Zip_New, OSC_Folder)
Sleep, 1000 ; Wait for unzip to finish

while (WinExist("ahk_pid" PID))
	Sleep, 10

SplitPath, OSC_Zip_New, vvName, vAppDir, vvEXT, vvNNE, vvDrive
OSC_Unzip := vAppDir . "\" . vvNNE
SplitPath, OSC_Folder,,vResourcesDir
Loop, Files, % OSC_Unzip "\*.*", F
	FileMove, % A_LoopFileFullPath, % vAppDir, 1
Loop, Files, % OSC_Unzip "\*.*", RD
	FileMoveDir, % A_LoopFileFullPath, % vResourcesDir "\app", 1
FileGetSize, vSize, % OSC_Unzip
If (vSize = 0)
	FileRemoveDir, % OSC_Unzip
If (!FileExist(OSC_Unzip))
	FileDelete, % OSC_Zip_New
Notify().AddWindow("has been Installed", {Title:vName, Font:"Sans Serif", TitleFont:"Sans Serif", Icon:userOSCPath "\open-stage-control.exe", Animate:"Right, Slide", ShowDelay:100, IconSize:64, TitleSize:14, Size:20, Radius:26, Time:2500, Background:"0x2C323A", Color:"0xD8DFE9", TitleColor:"0xD8DFE9"})
Run, % userOSCPath . "\open-stage-control.exe"
DetectHiddenWindows, Off

Sleep, 2500																; Give enough time for script to finish before exiting
ExitApp

ConvertResponseBody(oHTTP){
	Bytes := oHTTP.Responsebody ; Responsebody has an array of bytes. Single Characters
	Loop, % oHTTP.GetResponseHeader("Content-Length") ; Loop over Responsebody 1 byte (1 single character) at a time
		Text .= Chr(bytes[A_Index-1])
	return Text
}

BinArr_ToString(BinArr, Encoding := "UTF-8") {
	oADO := ComObjCreate("ADODB.Stream")

	oADO.Type := 1 ; adTypeBinary
	oADO.Mode := 3 ; adModeReadWrite
	oADO.Open
	oADO.Write(BinArr)

	oADO.Position := 0
	oADO.Type := 2 ; adTypeText
	oADO.Charset := Encoding
	return oADO.ReadText, oADO.Close
}

FileGetInfo( lptstrFilename) {
	List := "Comments InternalName ProductName CompanyName LegalCopyright ProductVersion"
	. " FileDescription LegalTrademarks PrivateBuild FileVersion OriginalFilename SpecialBuild"
	dwLen := DllCall("Version.dll\GetFileVersionInfoSize", "Str", lptstrFilename, "Ptr", 0)
	dwLen := VarSetCapacity( lpData, dwLen + A_PtrSize)
	DllCall("Version.dll\GetFileVersionInfo", "Str", lptstrFilename, "UInt", 0, "UInt", dwLen, "Ptr", &lpData)
	DllCall("Version.dll\VerQueryValue", "Ptr", &lpData, "Str", "\VarFileInfo\Translation", "PtrP", lplpBuffer, "PtrP", puLen )
	sLangCP := Format("{:04X}{:04X}", NumGet(lplpBuffer+0, "UShort"), NumGet(lplpBuffer+2, "UShort"))
	i := {}
	Loop, Parse, % List, %A_Space%
		DllCall("Version.dll\VerQueryValue", "Ptr", &lpData, "Str", "\StringFileInfo\" sLangCp "\" A_LoopField, "PtrP", lplpBuffer, "PtrP", puLen )
	? i[A_LoopField] := StrGet(lplpBuffer, puLen) : ""
	return i
}

updateSettingsIni() {
	global userOSCPath
	IniWrite, % userOSCPath, % A_ScriptDir "\settings.ini", settings, oscPath
}

checkUserOSCPath(byRef path) {
	pathExists := FileExist(path) == "D" ? true : false
	if (!pathExists) {
		MsgBox, 48, Error, OSC path not found. Please set a valid directory for your OSC install.
		InputBox, path, Set OSC Directory, Set a valid OSC Directory for the install
		If (ErrorLevel)
			ExitApp
		FileCreateDir, % Path
		If (ErrorLevel) {
			MsgBox, 48, Error, Could not create the path you listed, please run as admin or manually created the path and restart.
			ExitApp
		}
		checkUserOSCPath(path)
	}
	OnExit("updateSettingsIni")
	return path
}
; sZip = the fullpath of the zip file, sUnz the folder to contain the extracted files
Unzip(sZip, sUnz)	{
	FileCreateDir, %sUnz%
	psh := ComObjCreate("Shell.Application")
	psh.Namespace( sUnz ).CopyHere( psh.Namespace( sZip ).items, 4|16 )
}
