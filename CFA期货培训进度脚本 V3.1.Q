[General]
SyntaxVersion=2
BeginHotkey=32
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=27
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=d6a8fcb6-95af-43ff-8ba2-4c249c63f120
Description=CFA期货培训进度脚本 V3.1
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
//LogStart "d:\log.txt"
Hwnd = Plugin.Window.Find(0, "中国期货业协会网上后续培训系统-正式学习 - Mozilla Firefox")
Call Plugin.SysEx.Speed(Hwnd, 200)
Call Plugin.Window.Active(Hwnd)
Delay 500
Call Plugin.Window.Top(Hwnd, 0)
Delay 500
Call Plugin.Window.Size(Hwnd, 1200, 964)
Delay 500
Call Plugin.Window.Move(Hwnd, 0, 0)
Delay 500
Call Plugin.SysEx.SetInputMethod(Hwnd, "中文(简体)  - 美式键盘 ")
Delay 500
sRect = Plugin.Window.GetClientRect(Hwnd)
//sRect=114|0|1130|768
Delay 500
a=Split(sRect,"|")
Ux = CLng(a(0)) : Uy = CLng(a(1))//左上角X Y坐标
Dx = CLng(a(2)) : Dy = CLng(a(3)) //右下角X Y坐标
//设置指定窗口(Hwnd)加速到100倍


For timejs = 500 To 500 * 2 * 60 * 60 * 3 Step 500
    Delay 500*1
	Call nextkc()
Next


Sub playkc()
XY = Plugin.Pic.FindPic(ux+433, uy+359, ux+1091, uy+886, "d:\auto\play.bmp", 0, 0.5)
iZB = InStr(XY, "|")
X = CLng(Left(XY, iZB - 1))
Y = CLng(Right(XY, Len(XY) - iZB))
If x <= 0 and y <= 0 Then 
    Delay  kctime = Lib.算法.随机数字串(3)
else 
	//找到播放按钮 //点击播放
	MoveTo x+10, y+10
	LeftDown 1
	Delay 250
	LeftUp 1

	MoveTo 699, 607
	Delay 500
	LeftClick 1
	Delay 1
	LeftUp 1
	Delay 500
	LeftDown 1
	Delay 1
	LeftClick 1
	Delay 1
	LeftUp 1
	Delay 500
	
	For jsd = 0 To 20 Step 1
   	
   	  For jsd1 = 0 To 20 Step 1
   	   	kctime = Lib.算法.随机数字串(3)
		KeyDown "Right", 1
		Delay 1
		Delay kctime
		KeyUp "Right", 1
		Delay 3
		Delay 453
		 
		Delay 500*2*6
	   	Call nextkc()
	  Next
	Next
	
End If
End Sub

Sub nextkc()
XY = Plugin.Pic.FindPic(ux+1091, uy+417, ux+1170, uy+663, "d:\auto\next.bmp", 0, 0.5)
iZB = InStr(XY, "|")
X = CLng(Left(XY, iZB - 1))
Y = CLng(Right(XY, Len(XY) - iZB))
If x <= 0 and y <= 0 Then 
    //找不到下一页的情况
    XY2 = Plugin.Pic.FindPic(ux+448, uy+529, ux+508, uy+764, "d:\auto\A.bmp", 0, 0.5)
    iZB = InStr(XY2, "|")
	X2 = CLng(Left(XY2, iZB - 1))
	Y2 = CLng(Right(XY2, Len(XY2) - iZB))
	If x2 > 0 and y2 > 0 Then 
	   Call findan()
	   Else 
	   Call playkc()
	End If
else 
	//找到下一页按钮
	MsgBox "找到下一页"
	MoveTo ux+1139, uy+543
	LeftDown 1
	Delay 250
	LeftUp 1
	//点击下一页
	Delay 1000
	Call playkc()
End If
End Sub


Sub findtitle()
MoveTo ux+360, uy+304
Delay 500
LeftDown 1
MoveTo ux+1020, uy+308
Delay 500
LeftUp 1
Delay 1
LeftUp 1
Delay 500
KeyDown "Ctrl", 1
Delay 1
Delay 342
KeyDown "C", 1
Delay 1
Delay 215
KeyUp "C", 1
Delay 1
Delay 24
KeyUp "Ctrl", 1
Delay 4
kctitle = Plugin.Sys.GetCLB()
//MsgBox kctitle

End Sub

Sub findan()
	Call findtitle()
	Call Lib.系统.结束进程("excel.exe")
	Call Plugin.Office.OpenXls("D:\auto\CFA.xls")
    	For testy = 1 To 20 Step 1 //ym页码搜索
			//向Excel文档工作表(1)单元格(1, 1)读出内容存入(Text)变量
			Text = Plugin.Office.ReadXls(testy, 1, 1)
			If text = kctitle Then 
    			ym = testy
    			//MsgBox "该课程答案在" & ym & "页"
    			
    			
    			For testx = 1 To 40 Step 1
    				Text2 = Plugin.Office.ReadXls(ym, testx, 1)//题目标记
    				If text2 = "" Then 
    					TMBJ = testx - 1
    					Exit For
    				End If
    			Next
    			Exit For
    		End If
    	Next
   
    a = InputBox("请输入考试起始单元：") 
    For b = 1 To 30 Step 1
    	If a = left(Plugin.Office.ReadXls(ym, b, 1), 1) Then 
    		TM = b
    		Exit for
    	End If
    Next 
 
 	Delay 500
 	MsgBox tm&tmbj
    //For TM  To TMBJ Step 1//TM 题目自动加1,最多循环到50题
    While TM <= TMBJ
		For XX = 2 To 8 Step 1//XM选项自动跳最多循环到8个选项
    		daan = Plugin.Office.ReadXls(ym, TM, XX)
    		
    		XY4 = Plugin.Pic.FindPic(ux+448, uy+529, ux+508, uy+764, "d:\auto\A.bmp", 0, 0.5)
    		iZB = InStr(XY4, "|")
			X4 = CLng(Left(XY4, iZB - 1))
			Y4 = CLng(Right(XY4, Len(XY4) - iZB))
			If x2 > 0 and y2 > 0 Then 
			 
			 Select Case daan
				Case "A"
					ABCDEF="A"
					Call daclick(ABCDEF)
				Case "B"
					ABCDEF="B"
		   			Call daclick(ABCDEF)
				Case "C"
					ABCDEF="C"
		   			Call daclick(ABCDEF)
				Case "D"
					ABCDEF="D"
		   			Call daclick(ABCDEF)
				Case "E"
					ABCDEF="E"
		   			Call daclick(ABCDEF)
				Case "F"
					ABCDEF="F"
		   			Call daclick(ABCDEF)
		   		Case "G"
		   			ABCDEF="G"
		   			Call daclick(ABCDEF)
		   		Case ""
		   			Call COMMIT()//提交答案
		   			Exit For
		   		End Select
			Else 
			Exit Sub
			Call nextkc()
			 
			End If
		    
        Next
        TM=TM+1
    Wend

	Plugin.Office.CloseXls 
	
	End Sub
	
	
Sub daclick(ABCDEF)
		Select Case ABCDEF
			Case "A"
			    //MsgBox "A"
			    XY = Plugin.Pic.FindPic(ux+448, uy+529, ux+508, uy+764, "d:\auto\A.bmp", 0, 0.5)
    			iZB = InStr(XY, "|")
				X2 = CLng(Left(XY, iZB - 1))
				Y2 = CLng(Right(XY, Len(XY) - iZB))
				If x2 > 0 and y2 > 0 Then 
				    MoveTo X2, Y2 
					LeftDown 1
					Delay 250
					LeftUp 1
					Delay 250
                    //点击答案选项
				End If
			Case "B"
			    //MsgBox "B"
				XY = Plugin.Pic.FindPic(ux+448, uy+529, ux+508, uy+764, "d:\auto\B.bmp", 0, 0.5)
    			iZB = InStr(XY, "|")
				X2 = CLng(Left(XY, iZB - 1))
				Y2 = CLng(Right(XY, Len(XY) - iZB))
				If x2 > 0 and y2 > 0 Then 
				    MoveTo X2, Y2 
					LeftDown 1
					Delay 250
					LeftUp 1
					Delay 250
                    //点击答案选项
				End If
			Case "C"
			    //MsgBox "C"
				XY = Plugin.Pic.FindPic(ux+448, uy+529, ux+508, uy+764, "d:\auto\C.bmp", 0, 0.5)
    			iZB = InStr(XY, "|")
				X2 = CLng(Left(XY, iZB - 1))
				Y2 = CLng(Right(XY, Len(XY) - iZB))
				If x2 > 0 and y2 > 0 Then 
				    MoveTo X2, Y2 
					LeftDown 1
					Delay 250
					LeftUp 1
					Delay 250
                    //点击答案选项
				End If
			Case "D"
				XY = Plugin.Pic.FindPic(ux+448, uy+529, ux+508, uy+764, "d:\auto\D.bmp", 0, 0.5)
    			iZB = InStr(XY, "|")
				X2 = CLng(Left(XY, iZB - 1))
				Y2 = CLng(Right(XY, Len(XY) - iZB))
				If x2 > 0 and y2 > 0 Then 
				    MoveTo X2, Y2 
					LeftDown 1
					Delay 250
					LeftUp 1
					Delay 250
                    //点击答案选项
				End If
			Case "E"
				XY = Plugin.Pic.FindPic(ux+448, uy+529, ux+508, uy+764, "d:\auto\E.bmp", 0, 0.5)
    			iZB = InStr(XY, "|")
				X2 = CLng(Left(XY, iZB - 1))
				Y2 = CLng(Right(XY, Len(XY) - iZB))
				If x2 > 0 and y2 > 0 Then 
				    MoveTo X2, Y2 
					LeftDown 1
					Delay 250
					LeftUp 1
					Delay 250
                    //点击答案选项
				End If
			Case "F"
				XY = Plugin.Pic.FindPic(ux+448, uy+529, ux+508, uy+764, "d:\auto\F.bmp", 0, 0.5)
    			iZB = InStr(XY, "|")
				X2 = CLng(Left(XY, iZB - 1))
				Y2 = CLng(Right(XY, Len(XY) - iZB))
				If x2 > 0 and y2 > 0 Then 
				    MoveTo X2, Y2 
					LeftDown 1
					Delay 250
					LeftUp 1
					Delay 250
                    //点击答案选项
				End If
			Case "G"
				XY = Plugin.Pic.FindPic(ux+448, uy+529, ux+508, uy+764, "d:\auto\G.bmp", 0, 0.5)
    			iZB = InStr(XY, "|")
				X2 = CLng(Left(XY, iZB - 1))
				Y2 = CLng(Right(XY, Len(XY) - iZB))
				If x2 > 0 and y2 > 0 Then 
				    MoveTo X2, Y2 
					LeftDown 1
					Delay 250
					LeftUp 1
					Delay 250
                    //点击答案选项
				End If
		end Select
End Sub
	
Sub COMMIT()
	XY3 = Plugin.Color.FindColor(700, 620, 820, 793, "eaa227", 0, 1.0)
    iZB = InStr(XY3, "|")
	X3 = CLng(Left(XY3, iZB - 1))
	Y3 = CLng(Right(XY3, Len(XY3) - iZB))
	If x3 > 0 and y3 > 0 Then 
	    MoveTo X3 + 5, Y3 + 5 
		LeftDown 1
		Delay 250
		LeftUp 1
		//点击提交
    End If
    Delay 1000
	//CALL nextkc() //提交答案后随时检查是否可以跳到下一页,避免循环答题.
	End Sub
//LogStop
