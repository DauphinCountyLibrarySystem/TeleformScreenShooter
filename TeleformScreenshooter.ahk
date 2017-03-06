/*

	Name:		TeleformScreenshooter
	Version:	1.2
	Author:		Lucas Bodnyk

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
IniRead, bClosedClean, TeleformScreenshooter.ini, Log, bClosedClean, 0
IniRead, sEmailAs, TeleformScreenshooter.ini, General, sEmailAs, %A_Space%
IniRead, sPassword, TeleformScreenshooter.ini, General, sPassword, %A_Space%
IniRead, sRecipient, TeleformScreenshooter.ini, General, sRecipient, %A_Space%
IniRead, sScreenshotFilename, TeleformScreenshooter.ini, General, sScreenshotFilename, %A_Space%
IniRead, sSubject, TeleformScreenshooter.ini, General, sSubject, %A_Space%
IniRead, sBody, TeleformScreenshooter.ini, General, sBody, %A_Space%

Log("##        bClosedClean = "bClosedClean)
Log("##            sEmailAs = "sEmailAs)
Log("##           sPassword = "sPassword)
Log("##          sRecipient = "sRecipient)
Log("## sScreenshotFilename = "sScreenshotFilename)
Log("##            sSubject = "sSubject)
Log("##               sBody = "sBody)


If (bClosedClean = 0) {
	Log("!! It is likely that TeleformScreenshooter was terminated without warning.")
	}
If (sEmailAs = "") {
	Log("-- No username supplied. I won't be able to send an email without a username. Quitting.")
	ExitApp
	}
If (sPassword = "") {
	Log("-- No password supplied. I won't be able to send an email without a password. Quitting.")
	ExitApp
	}

IniWrite, 0, TeleformScreenshooter.ini, Log, bClosedClean
	
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
RunWait, MiniCap.exe -save "$appdir$%sScreenshotFilename%.jpg" -capturedesktop -exit,, Hide

EmailTheScreenshot:
Log("-- Emailing the file to %sRecipient%`,")
	RunWait, %A_WorkingDir%\mailsend1.18.exe -to %sRecipient% -from %sEmailAs% -ssl -smtp smtp.gmail.com -port 465 -sub "%sSubject%" -M "%sBody%" +cc +bc -q -auth-plain -user "%sEmailAs%" -pass "%sPassword%" -attach "%sScreenshotFilename%.jpg",, Hide
	
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
		Log("-- Exiting cleanly`, bClosedClean should be 1.")
		IniWrite, 1, TeleformScreenshooter.ini, Log, bClosedClean
	}
	if ExitReason in Menu
    {
        MsgBox, 4, , This takes screenshots for Allison.`nAre you sure you want to exit?
        IfMsgBox, No
            return 1  ; OnExit functions must return non-zero to prevent exit.
		IniWrite, 1, TeleformScreenshooter.ini, Log, bClosedClean
		Log("-- User is exiting TeleformScreenshooter`, dying now.")
    }
	if ExitReason in Logoff,Shutdown
	{
		IniWrite, 1, TeleformScreenshooter.ini, Log, bClosedClean
		Log("-- System logoff or shutdown in process`, dying now.")
	}
		if ExitReason in Close
	{
		IniWrite, 1, TeleformScreenshooter.ini, Log, bClosedClean
		Log("!! The system issued a WM_CLOSE or WM_QUIT`, or some other unusual termination is taking place`, dying now.")
	}
		if ExitReason not in Close,Exit,Logoff,Menu,Shutdown
	{
		IniWrite, 1, TeleformScreenshooter.ini, Log, bClosedClean
		Log("!! I am closing unusually`, with ExitReason: %ExitReason%`, dying now.")
	}
    ; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}