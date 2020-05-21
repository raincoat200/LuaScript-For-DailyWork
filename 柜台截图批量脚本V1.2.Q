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
MacroID=a283e995-7c75-48d4-b78d-52ddb8b6a3fa
Description=柜台截图批量脚本V1.2
Enable=0
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
   h=135
   //定义循环次数
Call Plugin.Office.OpenXls("E:\book.xls")
   Delay 400
   i = 1
   MoveTo 288, 1033
   Delay 300
   LeftClick 1
   //打开柜台
Do while i<h 
   Text1 = Plugin.Office.ReadXls(2, i + 1, 2)
   Text2 = Plugin.Office.ReadXls(2, i + 1, 3)
   Text3 = Plugin.Office.ReadXls(2, i + 1, 1)
   MoveTo 343, 113
   Delay 300
   LeftClick 1
   Delay 300
   KeyPress "BackSpace", 8
   Delay 300
   SayString Text1
   KeyDown "Enter", 1
   
   If text1 is null Then 
   MessageBox "Work finished!"
   Exit Do
   End If
   
   Delay 900
   KeyUp "Enter", 1
   Delay 500
   KeyDown "Enter", 1
   Delay 500
   KeyUp "Enter", 1
   Delay 500
   //取数完成
   KeyPress "Print Screen", 1
   Delay 400
   Call Lib.API.运行程序("C:\WINDOWS\system32\mspaint.exe")
   Delay 400
   Hwnd = Plugin.Window.Find("panint", "未命名 - 画图")
   Call Plugin.Window.Max(Hwnd)
   Delay 400
   Call Lib.键盘.键盘组合键("Ctrl + V", 0)
   Delay 400
   Call Lib.键盘.键盘组合键("Ctrl + S", 0)
   Delay 400
   SayString Text3&".jpg"
   Delay 400
   KeyDown "Enter", 1
   Delay 172
   Call Lib.API.激活窗口并置前(Hwnd)
   Call Lib.键盘.键盘组合键("ALT + F4", 0)
      
   If i>=h then
   Exit Do
   End If
   i = i + 1
Loop
MessageBox "Work finished!"
Plugin.Office.CloseXls
