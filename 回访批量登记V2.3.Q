[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=118
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=62f32bdf-5b16-422c-8362-ad4442a1a04e
Description=回访登记V2.3
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
Sub init()
	Hwnd = Plugin.Window.Find("MozillaWindowClass", "Mozilla Firefox")
    Call Plugin.Window.Show(Hwnd)
    Call Plugin.Window.Restore(Hwnd)
    Call Plugin.Window.Size(Hwnd, 1034,974)
    Call Plugin.Window.Restore(Hwnd)
    Delay 500
    Call Plugin.Window.Active(Hwnd)
    Delay 500
    //Call Plugin.Window.Top(Hwnd, 0)
    //Delay 500
    Call Plugin.Window.Move(Hwnd, 0, 0)
    Delay 500
End Sub

Sub main()
	MoveTo 955,263 //资金账号打叉
		Delay 1000
	LeftClick 1
		Delay 1000
	
	Call Plugin.Sys.SetCLB(zjzh)
	KeyDown 17, 1
	KeyPress 86, 1
	KeyUp 17, 1
	TracePrint zjzh &"  "&now()
	Delay 1000
	
	KeyPress "Enter", 1
	Delay 2000
	MoveTo 394,371
	Do
		
		XY = Plugin.Bkgnd.FindColorBlock(Hwnd, 230, 352, 396, 398, "F6E8DF", 5, 5, 0, 0.9)
		iZB = InStr(XY, "|")
		X = CLng(Left(XY, iZB - 1))
		Y = CLng(Right(XY, Len(XY) - iZB))
		If X>0 AND Y>0 then
		//If Plugin.Color.GetPixelColor(366,372,0)="F6E8DF" then
			//MoveTo 338, 371//第一条记录
			MoveTo X, Y  //第一条记录
			LeftClick 1
				Delay 1000
			MoveTo 315, 325  //回访登记按钮
				Delay 1000
			LeftClick 1
				Delay 500
			Exit Do
		Else 
			TracePrint "第一条记录颜色不对"
			Delay 500 
		End If
	Loop

	Do
		Delay 2000
		XY = Plugin.Color.FindColorBlock(72, 108, 524, 127, "EEDECF", 10, 5, 0, 0.9)
		iZB = InStr(XY, "|")
		X = CLng(Left(XY, iZB - 1))
		Y = CLng(Right(XY, Len(XY) - iZB))
		TracePrint X & Y
		If X > 0 AND Y > 0 Then 
			Exit Do
		End If
	loop

	MoveTo 144, 280//回访时间
	//MoveTo 144, 317//回访时间
	Delay 500
	LeftClick 1
	Delay 800
	
	Call Plugin.Sys.SetCLB(jtrq1)
	KeyDown 17, 1
	KeyPress 86, 1
	KeyUp 17, 1
	//TracePrint jtrq1
	Delay 800
	
	Select Case xs
		Case "现场"
			//do nothing
		Case "电话"
			MoveTo 179, 302
			//MoveTo 179, 343
				Delay 500
			LeftClick 1
				Delay 500
			MoveTo 160, 351
			//MoveTo 160, 385
				Delay 500
			LeftClick 1
				Delay 500
		Case "短信"
			MoveTo 179, 343
				Delay 500
			LeftClick 1
				Delay 500
			MoveTo 156, 473
				Delay 500
			LeftClick 1
				Delay 500
	End Select
	
	//MoveTo 174, 400
	MoveTo 174, 354
		Delay 500
	LeftClick 1
		Delay 800
	Call Plugin.Sys.SetCLB(nr)
	KeyDown 17, 1
	KeyPress 86, 1
	KeyUp 17, 1

	MoveTo 898,945 
		Delay 500
	LeftClick 1
		Delay 3000
End Sub

Dim zjzh, xm, jtrq,jtrq1, xs, nr
Dim i, j, maxnum
Call Plugin.Msg.HideScrTXT()
Call Plugin.Msg.ShowScrTXT(0, 0, 1024, 768, "确保已打开待回访客户明细页", "0000FF")
Call init()
Call Lib.系统.结束进程("excel.exe")
Call Plugin.Office.OpenXls("C:\Users\Administrator\Desktop\客户登记.xls")
//Call Plugin.Office.OpenXls("C:\客户登记.xls")

i=0
DO 
	For j = 1 To 5
		If IsNull(Plugin.Office.ReadXls(1, i + 2, j)) OR Plugin.Office.ReadXls(1, i + 2, j)="" Then 
			maxnum = i
			Call Plugin.Msg.ShowScrTXT(0, 0, 1024, 768, "合计" & maxnum & "条记录", "0000FF")
			TracePrint maxnum
			Plugin.Office.CloseXls 
			Exit do
		End If
	Next
	i=i+1
LOOP

For i = 2 To maxnum+1
	Call Plugin.Office.OpenXls("C:\Users\Administrator\Desktop\客户登记.xls")	
	if len(Plugin.Office.ReadXls(1, i, 6)) = 0
		zjzh = Plugin.Office.ReadXls(1, i, 1)
		xm = Plugin.Office.ReadXls(1, i, 2)
		jtrq = Plugin.Office.ReadXls(1, i, 3)
		//jtrq1 = mid(jtrq, 1, 4) & "-" & mid(jtrq, 5, 2) & "-" & mid(jtrq, 7, 2) & " 08:00"
		jtrq1 = mid(jtrq, 1, 4) & "-" & mid(jtrq, 5, 2) & "-" & mid(jtrq, 7, 2)
		//TracePrint jtrq1
		xs = Plugin.Office.ReadXls(1, i, 4)
		nr = Plugin.Office.ReadXls(1, i, 5)
		//Call Plugin.Msg.HideScrTXT()
		Call Plugin.Msg.ShowScrTXT(0, 0, 1024, 768, "第" & i-1 & "/"& maxnum &"条记录", "0000FF")
		Call Plugin.Msg.ShowScrTXT(0, 30, 1024, 768, zjzh, "0000FF")
		Call main()
		Call Plugin.Office.WriteXls(1, i, 6, "已登记")
	Else 
		Call Plugin.Msg.ShowScrTXT(0, 30, 1024, 768, "第"& i-1 &" 条已登记", "0000FF")
	End If
	Plugin.Office.CloseXls  
Next
