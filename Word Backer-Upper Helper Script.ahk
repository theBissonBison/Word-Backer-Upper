#NoEnv
SetWorkingDir %A_ScriptDir%
#singleinstance force
EnvGet, LocalAppData, LocalAppData
IniWrite, 1, %LocalAppData%\Word Backer-Upper\config.ini, section2, HelperScript
run, explorer "Word Backer-Upper.exe"
exitapp