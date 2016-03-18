;	Name:		TeleformScreenshooter
;	Version:	1.0
;	Author:		Lucas Bodnyk
;
;	This script draws code from the WinWait framework by berban on www.autohotkey.com, as well as some generic examples.
;	All variables "should" be prefixed with 'z'.
;
;
;	All User Startup is '\\<Machine_Name>\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup'.
;	I recommend placing a shortcut there, pointing to this, which should probably just go in the Sierra folder.

/*

Looks like the timer region of the screen is (758,508) to (932,567). I should just need to take a picture every minute and compare it to the previous picture. If they're the same, assume it crashed.
Still need to figure out the names of the controls if I'm going to manually shut it down and restart it.



*/

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
;#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
;#Persistent ; Keeps a script permanently running (that is, until the user closes it or ExitApp is encountered).
SplitPath, A_ScriptName, , , , ScriptBasename
StringReplace, AppTitle, ScriptBasename, _, %A_SPACE%, All
OnExit("ExitFunc") ; Register a function to be called on exit



;
;	BEGIN INITIALIZATION SECTION
;

Try {
	Log("")
	Log("   TeleformScreenshooter initializing for machine:" A_ComputerName)
} Catch	{
	MsgBox Testing TeleformScreenshooter.log failed! You probably need to check file permissions. I won't run without my log! Dying now.
	ExitApp
}
Try {
	IniWrite, 1, TeleformScreenshooter.ini, Test, zTest
	IniRead, zTest, TeleformScreenshooter.ini, Test, 0
	IniDelete, TeleformScreenshooter.ini, Test, zTest
} Catch {
	Log("!! Testing TeleformScreenshooter.ini failed! You probably need to check file permissions! I won't run without my ini! Dying now.")
	MsgBox Testing TeleformScreenshooter.ini failed! You probably need to check file permissions! I won't run without my ini! Dying now.
	ExitApp
}
IniRead, zClosedClean, TeleformScreenshooter.ini, Log, zClosedClean, 0
IniRead, zEmailAs, TeleformScreenshooter.ini, General, zEmailAs, %A_Space%
IniRead, zPassword, TeleformScreenshooter.ini, General, zPassword, %A_Space%
IniRead, zRecipient, TeleformScreenshooter.ini, General, zRecipient, %A_Space%

Log("## zClosedClean="zClosedClean)
Log("## zEmailAs="zEmailAs)
Log("## zPassword="zPassword)
Log("## zRecipient="Recipient)
If (zClosedClean = 0) {
	Log("!! It is likely that TeleformScreenshooter was terminated without warning.")
	}
If (zEmailAs = "") {
	Log("-- No username supplied. I won't be able to send an email without a username. Quitting.")
	ExitApp
	}
If (zPassword = "") {
	Log("-- No password supplied. I won't be able to send an email without a password. Quitting.")
	ExitApp
	}

IniWrite, 0, TeleformScreenshooter.ini, Log, zClosedClean
	
; DOWNLOAD NECESSARY FILES

Try {
	IfNotExist, %A_WorkingDir%\unzip.exe
		{
		Log("-- Trying to download unzip.exe...")
		URLDownloadToFile, http://stahlworks.com/dev/unzip.exe, %A_WorkingDir%\unzip.exe
		}
	IfNotExist, %A_WorkingDir%\mailsend1.18.exe
		{
		Log("-- Trying to download mailsend1.18.exe...")
		URLDownloadToFile, https://github.com/muquit/mailsend/releases/download/1.18/mailsend1.18.exe.zip, %A_WorkingDir%\mailsend1.18.exe.zip
		RunWait, %A_WorkingDir%\unzip.exe mailsend1.18.exe.zip,, Hide
		}
	IfNotExist, %A_WorkingDir%\MiniCap.exe
		{
		Log("-- Trying to download MiniCap.exe...")
		URLDownloadToFile, http://www.donationcoder.com/Software/Mouser/MiniCap/downloads/MiniCapPortable.zip, %A_WorkingDir%\MiniCapPortable.zip
		RunWait, %A_WorkingDir%\unzip.exe MiniCapPortable.zip,, Hide
		}
} Catch {
	Log("-- Necessary utilities not found and/or could not be acquired! Quitting.")
	ExitApp
}

Log("-- Initialization finished`, starting up... `(If I got this far`, it should mean I have all the necessary utilities`)")

TakeAScreenshot:
Log("-- Taking a screenshot of the desktop`,")
RunWait, MiniCap.exe -save "$appdir$screenshot.jpg" -capturedesktop -exit,, Hide

EmailTheScreenshot:
Log("-- Emailing the file to %zRecipient%`,")
	RunWait, %A_WorkingDir%\mailsend1.18.exe -to %zRecipient% -from %zEmailAs% -ssl -smtp smtp.gmail.com -port 465 -sub "III Teleforms screenshot" -M "Attached is a screenshot of the III Teleforms computer's desktop." +cc +bc -q -auth-plain -user "%zEmailAs%" -pass "%zPassword%" -attach "screenshot.jpg",, Hide
	
; functions to log and notify what's happening, courtesy of atnbueno
Log(Message, Type="1") ; Type=1 shows an info icon, Type=2 a warning one, and Type=3 an error one ; I'm not implementing this right now, since I already have custom markers everywhere.
{
	global ScriptBasename, AppTitle
	IfEqual, Type, 2
		Message = WW: %Message%
	IfEqual, Type, 3
		Message = EE: %Message%
	IfEqual, Message, 
		FileAppend, `n, %ScriptBasename%.log
	Else
		FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%.%A_MSec%%A_Tab%%Message%`n, %ScriptBasename%.log
	Sleep 50 ; Hopefully gives the filesystem time to write the file before logging again
	Type += 16
	;TrayTip, %AppTitle%, %Message%, , %Type% ; Useful for testing, but in production this will confuse my users.
	;SetTimer, HideTrayTip, 1000
	Return
	HideTrayTip:
	SetTimer, HideTrayTip, Off
	TrayTip
	Return
}
LogAndExit(message, Type=1)
{
	global ScriptBasename
	Log(message, Type)
	FileAppend, `n, %ScriptBasename%.log
	Sleep 1000
	ExitApp
}

 
ExitFunc(ExitReason, ExitCode)
{
    if ExitReason in Exit
	{
		Log("-- Exiting cleanly`, zClosedClean should be 1.")
		IniWrite, 1, TeleformScreenshooter.ini, Log, zClosedClean
	}
	if ExitReason in Menu
    {
        MsgBox, 4, , This takes screenshots for Allison.`nAre you sure you want to exit?
        IfMsgBox, No
            return 1  ; OnExit functions must return non-zero to prevent exit.
		IniWrite, 1, TeleformScreenshooter.ini, Log, zClosedClean
		Log("-- User is exiting TeleformScreenshooter`, dying now.")
    }
	if ExitReason in Logoff,Shutdown
	{
		IniWrite, 1, TeleformScreenshooter.ini, Log, zClosedClean
		Log("-- System logoff or shutdown in process`, dying now.")
	}
		if ExitReason in Close
	{
		IniWrite, 1, TeleformScreenshooter.ini, Log, zClosedClean
		Log("!! The system issued a WM_CLOSE or WM_QUIT`, or some other unusual termination is taking place`, dying now.")
	}
		if ExitReason not in Close,Exit,Logoff,Menu,Shutdown
	{
		IniWrite, 1, TeleformScreenshooter.ini, Log, zClosedClean
		Log("!! I am closing unusually`, with ExitReason: %ExitReason%`, dying now.")
	}
    ; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}