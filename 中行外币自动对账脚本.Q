[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=cd543b7f-93ef-4a73-8767-6bb0bbe86df9
Description=中行外币自动对账脚本
Enable=1
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]
Dim myhour, myvar, Length
Dim word
word = "XXXXX中行转帐接口机Ver3.5.3.8[2005.05.20]"
local_file = "D:\zhzz_th\bankcomm"
local_bat = "D:\zhzz_th\zgyh\Auto.BAT"

/*
word = "XXXXX中行转帐接口机Ver3.5.3.8[2005.05.20]"
local_file = "D:\zhzz_jfb\bankcomm"
local_bat = "D:\zhzz_jfb\bankcomm\Auto.BAT"
*/

Sub openexe()
	For 3
		Call Lib.系统.结束进程("cmd.exe")
		Call Lib.系统.结束进程("BankComm.exe")
		Delay 1000
	Next
	RunApp "c:\windows\system32\cmd.exe"
	Delay 1000
	SayString "d:"
	KeyPress "Enter", 1
	Delay 1000
	SayString "cd " & local_file
	KeyPress "Enter", 1
	Delay 1000
	SayString "BankComm.exe"
	KeyPress "Enter", 1
	Delay 100
	SayString "exit"
	KeyPress "Enter", 1
	Delay 50
End Sub

Sub init()
	For 5
		//TracePrint Plugin.Window.GetTextEx(Plugin.Window.Foreground(),0)
		If Plugin.Window.GetTextEx(Plugin.Window.Foreground(),0) <> word Then 
			Call Plugin.Window.Hide (Hwnd)
			Delay 500
			Call Plugin.Window.show (Hwnd)
			Delay 500
			Call Plugin.Window.active(Hwnd)
			Delay 500
			Call Plugin.Window.Top(Hwnd,0)
			Delay 500			
			Call Plugin.Window.Max(Hwnd)
			Delay 500
		Else 
			Exit For
		End If
	Next
End Sub

Sub getDZ()
	//UI Operation
	//TracePrint "start get dz progress"
	//ALT+F
	KeyDown 18, 1
	KeyPress 70, 1
	KeyUp 18, 1
	Delay 500
		
	//取银行端明细帐文件
	For 6
		KeyPress "Down", 1
		//Delay 100
	Next
	KeyPress "Enter", 1
	//回车确认提示
	Delay 3000
	KeyPress "Enter", 1
		
	//alt+f4
	KeyDown 18, 1
	KeyPress 115, 1
	KeyUp 18, 1	
	KeyPress "Enter", 1
	Delay 500
End Sub

Sub bat()
	RunApp "c:\windows\system32\cmd.exe"
	SayString "d:"
	KeyPress "Enter", 1
	Delay 1000
	SayString "cd " & local_file
	KeyPress "Enter", 1
	Delay 1000
	SayString local_bat
	KeyPress "Enter", 1
	Delay 1000
End Sub

Sub dz()
	For 100
		Hwnd = Plugin.Window.Find(0, word)
		//TracePrint  Hwnd	
		Length = Plugin.File.GetFileLength("O:\" & Replace(left(Plugin.Sys.GetDateTime, 10), "-", "") & ".txt")
		If Length <= 0 Then 
			If IsNull(Hwnd) or Hwnd = 0 Then 
				Call openexe()
				Delay 5000
			ElseIf Hwnd>0 then
			 	Call init()
				Call getDZ()
				Call bat()
				TracePrint "[收市] " & Plugin.Sys.GetDateTime
				Call netinit()	
			End If
		ElseIf Length > 0 Then
			Exit For	
		End If
		Delay 1000
	Next
End Sub

Sub netinit()
	RunApp "c:\windows\system32\cmd.exe"
	Delay 2000
	SayString "route add 26.0.0.0 mask 255.0.0.0 26.136.23.193"
	Delay 500
	KeyPress "Enter", 1
	Delay 2000
	SayString "exit"
	Delay 2000
	KeyPress "Enter", 1
End Sub


Call netinit()
Do 
	myhour = hour(now)
	myvar = minute(now)
	If myhour >= 16 and myhour < 23 and Weekday(now())<>7 and Weekday(now())<>1 Then 
		//周一周五16-22
		Call dz()
		Call Plugin.Msg.HideScrTXT()
		Delay 500
		Call Plugin.Msg.ShowScrTXT(0, 0, 1024, 768, "已完成对帐:" & now & " 文件大小:" & Length, "0000FF")
		Delay 500+1000*9
    ElseIf myhour >= 23 or Weekday(now())=7 or Weekday(now())=1 Then
    	//过了23点或者周末
    	Call Plugin.Msg.HideScrTXT()
    	Delay 500
    	Call Plugin.Msg.ShowScrTXT(0, 0, 1024, 768, "初始化成功:" & now, "0000FF")
		Delay 1000 * 30
    Else  
        Call Plugin.Msg.HideScrTXT()
        Delay 500
        Call Plugin.Msg.ShowScrTXT(0, 0, 1024, 768, "运行中:" & now, "0000FF")
        Delay 1000 * 30
    End If
Loop