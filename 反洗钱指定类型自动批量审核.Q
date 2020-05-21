[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=81
StopHotkeyMod=6
RunOnce=1
EnableWindow=
MacroID=d3ccf9bd-dd92-4f1c-9af4-22fd4af2817f
Description=反洗钱
Enable=1
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=1
QUI=Form1
[Relative]
SetupOCXFile=
[Comment]

[Attachment]
UEsDBBQAAgAIAFw36k5vRk3EjgAAAA4DAAAFABEAMS5ibXBVVA0AB4aMJV2GjCVdhowlXZ3RPQrCQBAF4KfCwC5oIQheIJgTGLRMF4k3EI/gz10VRKtUqVKlelbOprV6GeZ9xcLwii2PixmG2Xpyz9Iz90ywHt5fU/yNhaeFh2q6e1v4qPpdbCw0qqmvtdiqpr7OQqea+nqLvWrq+1qkqv/ViiMWyHJmG6oCxY7FnqrAoWZVUxU4nTligeudlxtVgR9QSwECFwsUAAIACABcN+pOb0ZNxI4AAAAOAwAABQAJAAAAAAAAACAAgIEAAAAAMS5ibXBVVAUAB4aMJV1QSwUGAAAAAAEAAQA8AAAAwgAAAAAA


[Script]
Sub init()
	Hwnd = Plugin.Window.Find("Chrome_WidgetWin_1", "反洗钱系统 - Google Chrome")
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

Sub find(char1, char2)
	//ctrl+F
	KeyDown 17, 1
	KeyPress 70, 1
	KeyUp 17, 1
	Delay 250
	SayString char1
	KeyPress "Enter", 1
	Delay 300
	XY = Plugin.Color.FindColor(0, 0, 947, 772, "3296FF", 0, 0.9)
	iZB = InStr(XY, "|")
	X = CLng(Left(XY, iZB - 1))
	Y = CLng(Right(XY, Len(XY) - iZB))
	If X > - 1  and Y > - 1  Then 
		If IsNumeric(char2) and len(char2)=1 Then 
			If char2 > 1 Then 
				For j = 1 To cint(char2)
					KeyPress "Enter", 1
					Delay 350
				Next	
			End If
			XY1 = Plugin.Color.FindColor(0, 0, 947, 772, "3296FF", 0, 0.9)
			iZB1 = InStr(XY1, "|")
			X1 = CLng(Left(XY1, iZB1 - 1))
			Y1 = CLng(Right(XY1, Len(XY1) - iZB1))
			If X1 > - 1  and Y1 > - 1  Then 
				MoveTo x1, y1
				LeftClick 1
			End If
			TracePrint "提交按钮:"&char1
		Else 
			MoveTo x, y
			//TracePrint x &" "& y
			Delay 150
			LeftClick 1
			Delay 300
			//CTRL+A DEL
			KeyDown 17, 1
			KeyPress 65, 1
			KeyUp 17, 1
			Delay 300
			KeyPress "Delete", 1
			Delay 150
			SayString char2
			TracePrint i&"条记录:"&char2
		End If 
		Delay 300
	Else 
		TracePrint "找不到:"& char1
	End If
End Sub

	Call init()
	intx = 591
	inty = 269
	
	Do
		MoveTo intx, inty
		LeftClick 2
		Delay 2000
		
		KeyDown 17, 1
		KeyPress 70, 1
		KeyUp 17, 1
		Delay 500
		
		SayString "证件过期"
		Delay 500
		
		FindColor 523,297,714,538,"3296FF",intX1,intY1
		If intX1 > 0 And intY1 > 0 Then 
			MoveTo 425, 588
			LeftClick 1
			Delay 500
			SayString "身份证件过期或失效风险等级上调一级"
			Delay 500
			
			
			FindPic 699,920,876,1013,"C:\temp\commit.bmp",0.8,intX1,intY1
			If intX1> 0 And intY1> 0 Then
				MoveTo intX1,intY1
				LeftClick 1
			Else 
				MoveTo 806, 960 //提交
				LeftClick 1
			End If
			Delay 1000
			
			Do
				FindPic 440,551,571,627,"C:\temp\1.bmp",0.8,intX1,intY1
				If intX1 > 0 And intY1 > 0 Then 
					MoveTo intX1,intY1
					LeftClick 1
					Delay 500
					js=js+1
					TracePrint js
					Exit Do
				Else 
					TracePrint "searching"
					Delay 1500
				End If
			loop
		Else 
			inty = inty + 25
			
			FindPic 795,103,877,174,"C:\temp\close.bmp",0.8,intX1,intY1
			If intX1> 0 And intY1> 0 Then
				MoveTo intX1,intY1
				LeftClick 1
				Delay 500
			Else 
				MoveTo 852, 135
				LeftClick 1
				Delay 500
				TracePrint "close 2"
			End If
			
			
			//close next
		End If
		If inty > 974 Then 
			MoveTo 198, 990
			LeftClick 1
			inty = 269
		End If
	Loop
	
	
	
