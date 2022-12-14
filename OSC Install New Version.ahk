; v1.3.13
#Requires Autohotkey v1.1.33+
#SingleInstance, Force ; Limit one running version of this script
SetBatchlines -1 ; run at maximum CPU utilization
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
#Include %A_ScriptDir%\Includes\Jxon.ahk
#Include %A_ScriptDir%\Includes\Notify.ahk
; --------------------------------------
checkAdminStatus()

global userOSCPath := checkUserOSCPath(oscPath)
global notificationSettings := {Title: ""
	, Font: "Sans Serif"
	, TitleFont: "Sans Serif"
	, Icon: A_ScriptDir "\assets\logo.png"
	, Animate: "Right, Slide"
	, ShowDelay: 100
	, IconSize: 64
	, TitleSize: 14
	, Size: 20
	, Radius: 26
	, Time: 999999
	, Background: "0x2C323A"
	, Color: "0xD8DFE9"
, TitleColor: "0xD8DFE9"}

currVersion := checkForCurrentVersion(userOSCPath)

Endpoint := "https://api.github.com/repos/jean-emmanuel/open-stage-control/releases/latest"
req := ComObjCreate("WinHttp.WinHttpRequest.5.1")
req.Open("GET", Endpoint, true)
req.Send()
req.WaitForResponse()

if (req.status != 200)
	errorTryAgain("ERROR", "There was an error with the request. Please try again later. `nError: " req.status)

res := req.ResponseBody
assets := JXON_Load(BinArr_ToString(res))
changelog := assets.body
latestVersion := SubStr(assets.tag_name, 2)

updateType := (changelog ~= "i)midi" | changelog ~= "i)electron") ? "major" : "minor"

if (currVersion == latestVersion) {
	installBetaOrPrev := upToDate(currVersion)
	if (installBetaOrPrev == false)
		ExitApp
	Endpoint := "https://api.github.com/repos/jean-emmanuel/open-stage-control/releases"
	req := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	req.Open("GET", Endpoint, true)
	req.Send()
	req.WaitForResponse()

	if (req.status != 200)
		errorTryAgain("ERROR", "There was an error with the request. Please try again later. `nError: " req.status)

	res := req.ResponseBody
	assets := JXON_Load(BinArr_ToString(res))
	assets := (installBetaOrPrev == "previous") ? assets[getPreviousVersions(assets)] : assets.1

	changelog := assets.body
	latestVersion := SubStr(assets.tag_name, 2)

	updateType := "major"
}

if (!FileExist(oscPath "\open-stage-control.exe"))
	updateType := "major"

fileRegex := updateType == "major" ? "i)win32-x64.zip" : "i)node.zip"

Loop % assets.assets.Length() {
	if (assets.assets[A_Index].name ~= fileRegex) {
		latestVersion := assets.assets[A_Index].name
		latestVersionURL := assets.assets[A_Index].browser_download_url
		Break
	}
}

if (!latestVersion)
	errorTryAgain("ERROR", "couldn't find OSC latest version for Windows")

MsgBox 0x41, OSC Installing New Version, % latestVersion ": was found`, would you like to install it?`nChanges Included:`n`n" changelog
IfMsgBox OK, {

} Else IfMsgBox Cancel, {
	ExitApp
}

notificationSettings.title := "found latest version"
Notify().AddWindow("downloading update now...", notificationSettings)

UrlDownloadToFile, % latestVersionURL, % A_ScriptDir "/" latestVersion

if WinExist("ahk_exe open-stage-control.exe") ; Close OSC before installing new version
	WinClose, % "ahk_exe open-stage-control.exe"

OSC_Zip := A_ScriptDir "\" latestVersion
OSC_Folder := updateType == "minor" ? userOSCPath "\resources\app" : userOSCPath
SplitPath, OSC_Zip, vName, vDir, vEXT, vNNE, vDrive
if (vEXT != "zip")
	errorTryAgain("ERROR", "There was an error with the downloading the update. Please try again later.")

FileDelete, % OSC_Folder "\*.*"
Loop, Files, % OSC_Folder "\*.*", D
{
	Path := OSC_Folder "\" A_LoopFileName
	if (Path != OSC_Folder "\_Archive") {
		FileRemoveDir, % Path, 1
	}
}

FileMove, % OSC_Zip, % OSC_Folder "\*.*", 1
OSC_Zip_New := OSC_Folder . "\" . vName
While(!FileExist(OSC_Zip_New))
	Sleep, 10

DetectHiddenWindows, On
Unzip(OSC_Zip_New, OSC_Folder)
Sleep, 1000 ; Wait for unzip to finish
SplitPath, OSC_Zip_New, vvName, vAppDir, vvEXT, vvNNE, vvDrive
OSC_Unzip := vAppDir . "\" . vvNNE
SplitPath, OSC_Folder,,vResourcesDir

destPath := updateType == "major" ? vResourcesDir . "\Open Stage Control" : vResourcesDir "\app"
Loop, Files, % OSC_Unzip "\*.*", F
	FileMove, % A_LoopFileFullPath, % destPath, 1
Loop, Files, % OSC_Unzip "\*.*", RD
	FileMoveDir, % A_LoopFileFullPath, % destPath, 1

FileGetSize, fileSize, % OSC_Unzip
If (fileSize = 0)
	FileRemoveDir, % OSC_Unzip
If (!FileExist(OSC_Unzip))
	FileDelete, % OSC_Zip_New

notificationSettings.title := vName
Notify().AddWindow("has been Installed", notificationSettings)

Run, % userOSCPath . "\open-stage-control.exe"
DetectHiddenWindows, Off

Sleep, 2500	; Give enough time for script to finish before exiting
ExitApp

checkForCurrentVersion(userOSCPath) {
	packageJSONPath := userOSCPath "\resources\app\package.json"
	if (!FileExist(packageJSONPath)) {
		currVersion := FileGetInfo(userOSCPath "\open-stage-control.exe").FileVersion
	}else {
		FileRead, packageJSON, % packageJSONPath
		packageJSON := JXON_Load(packageJSON)
		currVersion := packageJSON.version
	}
	return currVersion
}

getPreviousVersions(assets) {
	prevVersions := []
	for i,asset in assets
		prevVersions.push(asset.tag_name)

	return installPreviousVersion(prevVersions)
}

installPreviousVersion(prevVersions) {
	Instruction := "Select which to install:"
	Content := "in 30 seconds will default to the latest previous."
	Title := "Install Previous Version"
	MainIcon := 0xFFFD
	Flags := 0x818
	CustomButtons := []
	CustomButtons.Push([101, prevVersions.2])
	CustomButtons.Push([102, prevVersions.3])
	CustomButtons.Push([103, prevVersions.4])
	cButtons := CustomButtons.Length()
	VarSetCapacity(pButtons, 4 * cButtons + A_PtrSize * cButtons, 0)
	Loop %cButtons% {
		iButtonID := CustomButtons[A_Index][1]
		iButtonText := &(b%A_Index% := CustomButtons[A_Index][2])
		NumPut(iButtonID, pButtons, (4 + A_PtrSize) * (A_Index - 1), "Int")
		NumPut(iButtonText, pButtons, (4 + A_PtrSize) * A_Index - A_PtrSize, "Ptr")
	}
	DefaultButton := 101
	ipv_Callback := RegisterCallback("ipv_Callback", "Fast")
	CBData := {}
	CBData.Timeout := 30000 ; ms

	; TASKDIALOGCONFIG structure
	x64 := A_PtrSize == 8
	NumPut(VarSetCapacity(TDC, x64 ? 160 : 96, 0), TDC, 0, "UInt") ; cbSize
	NumPut(Flags, TDC, x64 ? 20 : 12, "Int") ; dwFlags
	NumPut(&Title, TDC, x64 ? 28 : 20, "Ptr") ; pszWindowTitle
	NumPut(MainIcon, TDC, x64 ? 36 : 24, "Ptr") ; pszMainIcon
	NumPut(&Instruction, TDC, x64 ? 44 : 28, "Ptr") ; pszMainInstruction
	NumPut(&Content, TDC, x64 ? 52 : 32, "Ptr") ; pszContent
	NumPut(cButtons, TDC, x64 ? 60 : 36, "UInt") ; cButtons
	NumPut(&pButtons, TDC, x64 ? 64 : 40, "Ptr") ; pButtons
	NumPut(DefaultButton, TDC, x64 ? 72 : 44, "Int") ; nDefaultButton
	NumPut(ipv_Callback, TDC, x64 ? 140 : 84, "Ptr") ; pfCallback
	NumPut(&CBData, TDC, x64 ? 148 : 88, "Ptr") ; lpCallbackData

	DllCall("Comctl32.dll\TaskDialogIndirect", "Ptr", &TDC
		, "Int*", Button := 0
		, "Int*", Radio := 0
	, "Int*", Checked := 0)

	DllCall("Kernel32.dll\GlobalFree", "Ptr", ipv_Callback)

	If (Button == 101 || Button == 2) { ; 1
		return 2
	} Else If (Button == 102) { ; 2
		return 3
	} Else If (Button == 103) { ; 3
		return 4
	}
}

ipv_Callback(hWnd, Notification, wParam, lParam, RefData) {
	Local CBData := Object(RefData)

	If (Notification == 4 && wParam > CBData.Timeout) {
		; TDM_CLICK_BUTTON := 0x466, IDCANCEL := 2
		PostMessage 0x466, 2, 0,, ahk_id %hWnd%
	}
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

FileGetInfo(lptstrFilename) {
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
	if (!FileExist(A_ScriptDir "\settings.ini"))
		IniWrite, % "", % A_ScriptDir "\settings.ini", settings, oscPath

	IniRead, path, % A_ScriptDir "\settings.ini", settings, oscPath
	pathExists := FileExist(path) == "D" ? true : false
	if (!pathExists) {
		MsgBox, 48, Error, OSC path not found. Please set a valid directory for your OSC install.
		InputBox, path, Set OSC Directory, Set a valid OSC Directory for the install
		If (ErrorLevel)
			ExitApp
		if (!FileExist(path))
			FileCreateDir, % path
		If (ErrorLevel) {
			MsgBox, 48, Error, Could not create the path you listed, try to manually create the path and restart.
			ExitApp
		}
		IniWrite, % path, % A_ScriptDir "\settings.ini", settings, oscPath
		checkUserOSCPath(path)
	}
	OnExit("updateSettingsIni")
	return path
}

Unzip(zipFullPath, outputDir) {
	FileCreateDir, %outputDir%
	psh := ComObjCreate("Shell.Application")
	psh.Namespace( outputDir ).CopyHere( psh.Namespace( zipFullPath ).items, 4|16 )
}

UpToDate(version) {
	Instruction := "You're currently up to date!"
	Content := "Your current version is: " version
	Title := "Checking for Updates"
	MainIcon := 0xFFFD
	Flags := 0x810
	CustomButtons := []
	CustomButtons.Push([101, "Check for Beta?"])
	CustomButtons.Push([102, "Install Previous Version(s)?"])
	CustomButtons.Push([103, "Exit"])
	cButtons := CustomButtons.Length()
	VarSetCapacity(pButtons, 4 * cButtons + A_PtrSize * cButtons, 0)
	Loop %cButtons% {
		iButtonID := CustomButtons[A_Index][1]
		iButtonText := &(b%A_Index% := CustomButtons[A_Index][2])
		NumPut(iButtonID, pButtons, (4 + A_PtrSize) * (A_Index - 1), "Int")
		NumPut(iButtonText, pButtons, (4 + A_PtrSize) * A_Index - A_PtrSize, "Ptr")
	}
	DefaultButton := 103
	TDCallback_UpToDate := RegisterCallback("TDCallback_UpToDate", "Fast")
	CBData := {}
	CBData.Timeout := 30000 ; ms

	; TASKDIALOGCONFIG structure
	x64 := A_PtrSize == 8
	NumPut(VarSetCapacity(TDC, x64 ? 160 : 96, 0), TDC, 0, "UInt") ; cbSize
	NumPut(Flags, TDC, x64 ? 20 : 12, "Int") ; dwFlags
	NumPut(&Title, TDC, x64 ? 28 : 20, "Ptr") ; pszWindowTitle
	NumPut(MainIcon, TDC, x64 ? 36 : 24, "Ptr") ; pszMainIcon
	NumPut(&Instruction, TDC, x64 ? 44 : 28, "Ptr") ; pszMainInstruction
	NumPut(&Content, TDC, x64 ? 52 : 32, "Ptr") ; pszContent
	NumPut(cButtons, TDC, x64 ? 60 : 36, "UInt") ; cButtons
	NumPut(&pButtons, TDC, x64 ? 64 : 40, "Ptr") ; pButtons
	NumPut(DefaultButton, TDC, x64 ? 72 : 44, "Int") ; nDefaultButton
	NumPut(TDCallback_UpToDate, TDC, x64 ? 140 : 84, "Ptr") ; pfCallback
	NumPut(&CBData, TDC, x64 ? 148 : 88, "Ptr") ; lpCallbackData

	DllCall("Comctl32.dll\TaskDialogIndirect", "Ptr", &TDC
		, "Int*", Button := 0
		, "Int*", Radio := 0
	, "Int*", Checked := 0)

	DllCall("Kernel32.dll\GlobalFree", "Ptr", TDCallback_UpToDate)

	If (Button == 101) { ; Check for Beta?
		return "beta"
	} Else If (Button == 102) { ; Install Previous Version(s)?
		return "previous"
	} Else If (Button == 103) { ; Exit
		return false
	} Else If (Button == 2) { ; Timeout
		return false
	}
}

TDCallback_UpToDate(hWnd, Notification, wParam, lParam, RefData) {
	Local CBData := Object(RefData)

	If (Notification == 4 && wParam > CBData.Timeout) {
		; TDM_CLICK_BUTTON := 0x466, IDCANCEL := 2
		PostMessage 0x466, 2, 0,, ahk_id %hWnd%
	}
}

checkAdminStatus() {
	if (!A_IsAdmin){ ;http://ahkscript.org/docs/Variables.htm#IsAdmin
		Instruction := "Restarting App..."
		Content := "You must run this script as an administrator"
		Title := "NOT ADMIN"
		MainIcon := 0xFFFC
		Flags := 0xE00
		Buttons := 0x20
		TDCallback := RegisterCallback("TDCallback", "Fast")
		CBData := {}
		CBData.Marquee := True
		CBData.Timeout := 3000 ; ms

		; TASKDIALOGCONFIG structure
		x64 := A_PtrSize == 8
		NumPut(VarSetCapacity(TDC, x64 ? 160 : 96, 0), TDC, 0, "UInt") ; cbSize
		NumPut(Flags, TDC, x64 ? 20 : 12, "Int") ; dwFlags
		NumPut(Buttons, TDC, x64 ? 24 : 16, "Int") ; dwCommonButtons
		NumPut(&Title, TDC, x64 ? 28 : 20, "Ptr") ; pszWindowTitle
		NumPut(MainIcon, TDC, x64 ? 36 : 24, "Ptr") ; pszMainIcon
		NumPut(&Instruction, TDC, x64 ? 44 : 28, "Ptr") ; pszMainInstruction
		NumPut(&Content, TDC, x64 ? 52 : 32, "Ptr") ; pszContent
		NumPut(TDCallback, TDC, x64 ? 140 : 84, "Ptr") ; pfCallback
		NumPut(&CBData, TDC, x64 ? 148 : 88, "Ptr") ; lpCallbackData

		DllCall("Comctl32.dll\TaskDialogIndirect", "Ptr", &TDC
			, "Int*", Button := 0
			, "Int*", Radio := 0
		, "Int*", Checked := 0)

		DllCall("Kernel32.dll\GlobalFree", "Ptr", TDCallback)

		If (Button == 8) { ; Close
			Run *RunAs "%A_ScriptFullPath%" ; Requires v1.0.92.01+
			ExitApp

		} Else If (Button == 2) { ; Timeout
			Run *RunAs "%A_ScriptFullPath%",, UseErrorLevel ; Requires v1.0.92.01+
			if(ErrorLevel)
				ExitApp
			ExitApp
		}
	}
}

errorTryAgain(title, message) {
	OnMessage(0x44, "OnReqStatusError")
	MsgBox 0x11, ERROR, % "There was an error with the request. Please try again later. `nError: " status
	OnMessage(0x44, "")

	IfMsgBox OK, {
		Reload
	} Else IfMsgBox Cancel, {
		ExitApp
	}
}

OnReqStatusError() {
	DetectHiddenWindows, On
	Process, Exist
	If (WinExist("ahk_class #32770 ahk_pid " . ErrorLevel)) {
		ControlSetText Button1, Try Again
		ControlSetText Button2, Exit
	}
}

TDCallback(hWnd, Notification, wParam, lParam, RefData) {
	Local CBData := Object(RefData)

	If (Notification == 4 && wParam > CBData.Timeout) {
		; TDM_CLICK_BUTTON := 0x466, IDCANCEL := 2
		PostMessage 0x466, 2, 0,, ahk_id %hWnd%
	}

	If (Notification == 0 && CBData.Marquee) {
		; TDM_SET_PROGRESS_BAR_MARQUEE
		DllCall("PostMessage", "Ptr", hWnd, "UInt", 0x46B, "UInt", 1, "UInt", 50)
	}
}