#Persistent
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#singleinstance force
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

EnvGet, AppData, AppData
EnvGet, LocalAppData, LocalAppData
EnvGet, USERPROFILE, USERPROFILE
MainClickedOn = 1

IniRead, HelperScript, %LocalAppData%\Word Backer-Upper\config.ini, section2, HelperScript
if HelperScript = 1
{
	MainClickedOn = 0
	HelperScript = 0
	IniWrite, 0, %LocalAppData%\Word Backer-Upper\config.ini, section2, HelperScript
}

IniRead, Genesis, %LocalAppData%\Word Backer-Upper\config.ini, Identity, Genesis
if (Genesis != 1)
{
	FileCreateDir, %LocalAppData%\Word Backer-Upper
	FileCreateDir, %USERPROFILE%\Documents\Word Backer-Upper
	CopyLoc = %USERPROFILE%\Documents\Word Backer-Upper
	CheckInt = 300000
	SaveInt = 30000
	BackupInt = 10
	BackupsEnabled = 1
	IniWrite, %CopyLoc%, %LocalAppData%\Word Backer-Upper\config.ini, section1, CopyLoc
	IniWrite, %CheckInt%, %LocalAppData%\Word Backer-Upper\config.ini, section1, CheckInt
	IniWrite, %SaveInt%, %LocalAppData%\Word Backer-Upper\config.ini, section1, SaveInt
	IniWrite, %BackupInt%, %LocalAppData%\Word Backer-Upper\config.ini, section1, BackupInt
	IniWrite, %BackupsEnabled%, %LocalAppData%\Word Backer-Upper\config.ini, section1, BackupsEnabled
	
	IniWrite, 1, %LocalAppData%\Word Backer-Upper\config.ini, Identity, Genesis
	MainClickedOn = 1
}

IniRead, CopyLoc, %LocalAppData%\Word Backer-Upper\config.ini, section1, CopyLoc
IniRead, CheckInt, %LocalAppData%\Word Backer-Upper\config.ini, section1, CheckInt
IniRead, SaveInt, %LocalAppData%\Word Backer-Upper\config.ini, section1, SaveInt
IniRead, BackupInt, %LocalAppData%\Word Backer-Upper\config.ini, section1, BackupInt
IniRead, BackupsEnabled, %LocalAppData%\Word Backer-Upper\config.ini, section1, BackupsEnabled

SettingsPart:
if MainClickedOn = 1
{
	editCheckInt := CheckInt/60000
	editSaveInt := SaveInt/60000
	editBackupInt := BackupInt
	editBackupsEnabled := BackupsEnabled
	
	if WinExist("ahk_id" hwndsettingsbutton_backupper)
	{
		Gui, %hwndsettingsbutton_backupper%:Destroy
	}
	Gui, New
	Gui, +hwndhwndsettingsbutton_backupper
	
	Gui, Font, s10, Segoe UI
	Gui, Add, Text, ,Backups folder:
	Gui, add, Edit, w300 vFSFback
	GuiControl,, FSFback, %CopyLoc%
	Gui, Add, Button,  w30 gBackupfolder, . . .
	
	Gui, Add, Text,,Backup the recovery file versions?`n(Can prevent backup file loss when word crashes):
	Gui, Add, CheckBox, veditBackupsEnabled checked%editBackupsEnabled%, Enabled
	
	Gui, Add, Text,,Interval to check for Word documents (min):
	Gui, add, Edit, veditCheckInt
	GuiControl,, editCheckInt, %editCheckInt%
	
	Gui, Add, Text,,Interval to save Word documents (min):
	Gui, add, Edit, veditSaveInt
	GuiControl,, editSaveInt, %editSaveInt%
	
	Gui, Add, Text,,Interval to backup Word recovery file:`n(Once per how many saves?)
	Gui, add, Edit, veditBackupInt
	GuiControl,, editBackupInt, %editBackupInt%
	
	Gui, Add, Button, x160 y350 gSettingsDefault, Revert to Defaults
	Gui, Add, Button, x290 y350 w50 gSettingsApply, OK
	
	Gui, Show,, Word Backer-Upper Settings
	return
}

MainPart:
Indexer = 1
SetTitleMatchMode, 2

Loop
{
	if WinExist("- Word") 
	{
		Loop
		{			
			if WinActive("- Word")
			{
				Send, ^s
				Indexer := Indexer + 1
				if ((Mod(Indexer, BackupInt) = 0 OR Indexer = 2) AND BackupsEnabled = 1)
				{
					FormatTime, CurrentTime , , M-d-yy HH.mm
					FileCopyDir, %AppData%\Microsoft\Word, %CopyLoc%\%CurrentTime% Word_Backup
				}
			}
			if !WinExist("- Word")
			{
				Break
			}
			sleep, %SaveInt%
		}
	}
	
	if ((Mod(A_Index, 30) = 0 OR Indexer > 1) AND BackupsEnabled = 1)
	{
		TimeCode = % A_NowUTC
		FormatTime, Today, %TimeCode%, M-d-yy
		
		;EnvAdd, TimeCode, -1, Days
		;FormatTime, Yesterday, %TimeCode%, M-d-yy
		
		Loop, Files, %CopyLoc%\*Word_Backup, D
		{
			if !(InStr(A_LoopFileName, Today)) ;OR InStr(A_LoopFileName, Yesterday))
			{
				FileRemoveDir, %A_LoopFileFullPath%, 1
			}
		}
	}
	Indexer = 1
	sleep, %CheckInt%
}
return


Backupfolder:
FileSelectFolder, FSFback,,3,Choose desktop folder:
if (FSFback = "")
{
	msgbox, Folder selection failed.
	FSFback := CopyLoc
}
GuiControl,, FSFback, %FSFback%
WinActivate, ahk_id %hwndsettingsbutton_backupper%
return

SettingsDefault:
CopyLoc := "B:\Backup\Word"
CheckInt = 300000
SaveInt = 30000
BackupInt = 10
BackupsEnabled = 1
IniWrite, %CopyLoc%, %LocalAppData%\Word Backer-Upper\config.ini, section1, CopyLoc
IniWrite, %CheckInt%, %LocalAppData%\Word Backer-Upper\config.ini, section1, CheckInt
IniWrite, %SaveInt%, %LocalAppData%\Word Backer-Upper\config.ini, section1, SaveInt
IniWrite, %BackupInt%, %LocalAppData%\Word Backer-Upper\config.ini, section1, BackupInt
IniWrite, %BackupsEnabled%, %LocalAppData%\Word Backer-Upper\config.ini, section1, BackupsEnabled
Gui, cancel
Gui, Destroy
goto, SettingsPart
return

SettingsApply:
CheckInt := editCheckInt*60000
SaveInt := editSaveInt*60000
BackupInt := editBackupInt
BackupsEnabled := editBackupsEnabled
CopyLoc := FSFback

IniWrite, %CopyLoc%, %LocalAppData%\Word Backer-Upper\config.ini, section1, CopyLoc
IniWrite, %CheckInt%, %LocalAppData%\Word Backer-Upper\config.ini, section1, CheckInt
IniWrite, %SaveInt%, %LocalAppData%\Word Backer-Upper\config.ini, section1, SaveInt
IniWrite, %BackupInt%, %LocalAppData%\Word Backer-Upper\config.ini, section1, BackupInt
IniWrite, %BackupsEnabled%, %LocalAppData%\Word Backer-Upper\config.ini, section1, BackupsEnabled
Gui, cancel
Gui, Destroy
goto, MainPart
return