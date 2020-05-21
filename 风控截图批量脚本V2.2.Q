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
MacroID=1f8c032a-369c-4d00-9c6b-f76f0ad59238
Description=风控截图批量脚本V2.2
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
//定义excel路径
Call Plugin.Office.OpenXls("c:\king.xls")
Call Plugin.File.CreateFolder("c:\temp")
Delay 800
yanshi = 200
//打开风控窗口
//MoveTo 334, 1029
//Delay 1000
//LeftClick 1

For i=1 to i+1
   Text1 = Plugin.Office.ReadXls(1, i + 1, 4)  //资金账号：  表1,2行，4列
   Text2 = Plugin.Office.ReadXls(1, i + 1, 5)  //姓名：      表1,2行，5列
   Text3 = Plugin.Office.ReadXls(1, i + 1, 1)  //序号：      表1,2行，1列
   Text4 = Plugin.Office.ReadXls(1, i + 1, 3)  //营业部名称： 表1,2行，3列
   Text5 = Plugin.Office.ReadXls(1, i + 1, 6)  //处理标记：   表1,2行，6列
   isfile = Lib.文件.判断文件夹是否存在("c:\temp\"&Text4)
   lujin = "c:\temp\" & Text4 & "\"
   wjname = lujin & Text3 & ".jpg"
   
   Text1 = Replace(Text1, "-", "")
   Text1 = Replace(Text1, "(", "")
   Text1 = Replace(Text1, ")", "")
   Text1 = Replace(Text1, " ", "")
   Text2 = Replace(Text2, "-", "")
   Text2 = Replace(Text2, "(", "")
   Text2 = Replace(Text2, ")", "")
   Text2 = Replace(Text2, " ", "")
   Text3 = Replace(Text3, "-", "")
   Text3 = Replace(Text3, "(", "")
   Text3 = Replace(Text3, ")", "")
   Text3 = Replace(Text3, " ", "")
   Text4 = Replace(Text4, "-", "")
   Text4 = Replace(Text4, "(", "")
   Text4 = Replace(Text4, ")", "")
   Text4 = Replace(Text4, " ", "")
     
   if isfile= false then 
   Call Plugin.File.CreateFolder("c:\temp\"&Text4)
   Else 
   End If
   
   If Text1 ="" or Text2 ="" or Text3 ="" Then 
   Exit For
   ElseIf Text5 = "" or Plugin.File.IsFileExist(wjname) = False Then
   //处理标记为空及未检测到文件放执行
   	
   //移动到客户风控等级查询坐标   
   MoveTo 976, 199
   Delay yanshi
   LeftDown 1
   Delay yanshi
   LeftUp 1
   Delay yanshi
   KeyPress "BackSpace", 12
   Delay yanshi
   SayString Text1
   KeyDown "Enter", 1
   Delay yanshi
   KeyUp "Enter", 1
   Delay 2000
   //窗口调好大小
   MoveTo 356, 332
   Delay yanshi
   LeftDown 1
   Delay 3
   LeftClick 1
   Delay 1
   LeftUp 1
   Delay 500
   RightClick 1
   RightClick 1
   Delay 85
   RightClick 1
   Delay 2
   RightClick 1
   MoveTo 393, 416
   Delay yanshi
   LeftDown 1
   Delay 2
   LeftClick 1
   Delay 1
   LeftUp 1
   MoveTo 622, 284
   Delay 500
   LeftDown 1
   Delay 2
   LeftDown 1
   MoveTo 382, 401
   Delay 500
   LeftUp 1
   MoveTo 940, 853
   Delay 500
   LeftDown 1
   Delay 2
   LeftDown 1
   MoveTo 1294, 836
   Delay 500
   LeftUp 1
   MoveTo 1189, 401
   Delay 500
   LeftDown 1
   Delay 2
   LeftClick 1
   //启动截屏
   KeyPress "Print Screen", 1
   Delay 400
   Call Lib.API.运行程序("C:\WINDOWS\system32\mspaint.exe")
   Delay yanshi
   Hwnd = Plugin.Window.Find("MSPaintApp", "无标题 - 画图")
   Call Plugin.Window.Top(Hwnd, 0)
   Delay yanshi
   Call Lib.键盘.键盘组合键("Ctrl + V", 0)
   Delay yanshi
   Call Lib.键盘.键盘组合键("Ctrl + S", 0)
   Delay yanshi
   //写文件名，存盘
   KeyDown 17, 1
   KeyPress 65, 1
   KeyUp 17, 1
   KeyPress "BackSpace", 4
   Delay yanshi
   SayString wjname
   Delay yanshi
   KeyDown "Enter", 1
   Delay yanshi
    //关闭截屏
   Call Lib.API.激活窗口并置前(Hwnd)
   Call Lib.键盘.键盘组合键("ALT + F4", 0)
   //写处理标记
   Call Plugin.Office.WriteXls(1, i + 1, 6, "已完成")
   //关闭当前风控子窗口   
   MoveTo 1287, 397
   Delay yanshi
   LeftDown 1
   Delay 1
   LeftClick 1
   
   End if
   Next
Plugin.Office.CloseXls 
MsgBox "批处理完成！"&i&"次"

