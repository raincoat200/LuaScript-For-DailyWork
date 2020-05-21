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
MacroID=b8226028-765d-49f4-bea6-94c2fe71bcda
Description=SAC培训进度脚本
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
Hwnd = Plugin.Window.Find(0, "在线学习 - Mozilla Firefox")
Call Plugin.Window.Active(Hwnd)
Delay 500
Call Plugin.Window.Top(Hwnd, 0)
Delay 500
Call Plugin.Window.Size(Hwnd, 1024, 960)
Delay 500
Call Plugin.SysEx.SetInputMethod(Hwnd, "中文(简体)  - 美式键盘 ")
Delay 500
sRect = Plugin.Window.GetClientRect(Hwnd)
//sRect=114|0|1130|768
Delay 500
a=Split(sRect,"|")
Ux = CLng(a(0)) : Uy = CLng(a(1))//左上角X Y坐标
Dx = CLng(a(2)) : Dy = CLng(a(3)) //右下角X Y坐标

zhangay=array("第一章","第二章","第三章","第四章","第五章","第六章","第七章","第八章","第九章","第十章","第十一章","第十二章","第十三章","第十四章","第十五章","第十六章","第十七章","第十八章","第十九章","第二十章","第二十一章","第二十二章","第二十三章","第二十四章","第二十五章","第二十六章") 
keay = array("讲课1", "讲课2", "讲课3", "讲课4", "讲课5", "讲课6", "讲课7", "讲课8", "讲课9", "讲课10")

/* 
Call FINDF4()
SayString "课后"
Delay 250
KeyPress "Enter", 1
Delay 500
endi=0
For i = 0 To 35 Step 1
    Select Case i
    case 0 
	   Call lastpage()
	case 12  //大于13章即翻页
		Call nextpage()
	case 25  //大于25章即翻页
		Call nextpage()
	End Select
	
	FindColor ux+15, uy+255, ux+228, dy-34, "78D838", intX, intY //查找测验页
    If intX > 0 And intY > 0 Then 
    endi = i - 1
    MsgBox endi
    Delay 500
    KeyPress "Enter", 1
    End If
Next
//查答案页
*/



For i = 0 To 30 Step 1
    Select Case i
    case 0 
	   Call lastpage()
	case 12  //大于13章即翻页
		Call nextpage()
	case 25  //大于25章即翻页
		Call nextpage()
	End Select
	
	Call FINDF4()
	SayString zhangay(i)
	Delay 250
	KeyPress "Enter", 1
	Delay 500
	FindColor ux+15, uy+255, ux+228, dy-34, "78D838", intX, intY //查找章节
	//MsgBox intX & "-" & intY	
	If intX > 0 And intY > 0 Then     
		MoveTo intX, intY
		LeftDown 1
		Delay 250
		LeftUp 1
		//单击章节
		For j=0 to 9  step 1
			Call FINDF4()
			SayString keay(j)
			Delay 250
			KeyPress "Enter", 1
			Delay 200
			FindColor ux + 15, uy + 255, ux + 228, dy - 34, "78D838", intX, intY//查找讲课
			If intX > 0 And intY > 0 Then 
				MoveTo intX+55, intY+16 //偏移定位点击
				LeftDown 1
				Delay 250
				LeftUp 1
				Delay 3000 //视频缓冲
				//单击讲课-播放
					
				Call FINDF4()
				SayString "我的观看轨迹"
				Delay 250
				KeyPress "Enter", 1
				Delay 500
				FindColor ux + 248, uy + 800, ux +350, uy + 840, "78D838", intX, intY //查找轨迹
				  If intX > 0 And intY > 0 Then 
					MoveTo intX, intY
    				LeftDown 1
					Delay 800 //点击轨迹查看延时
					LeftUp 1
						
					Call FINDF4()		
					SayString "100%"
					Delay 250
					KeyPress "Enter", 1
					Delay 500
					FindColor ux + 524, uy + 828, ux + 591, uy + 865, "78D838", intX, intY//查找100%进度
					  If intX > 0 And intY > 0 Then 
		    		   //找到100%进度
		    		  Else 
		    		  	ux1 = ux + 350
                        uy1 = ux + 392
                        For l = 0 To 120 Step 1
					  	MoveTo ux1, uy1				  	
					  	XY = Plugin.Pic.FindPic(Ux+256, Uy+386, Ux+973, Uy+785, "D:\finish.bmp", 0, 0.5)
					  	iZB = InStr(XY, "|")
					  	X = CLng(Left(XY, iZB - 1))
					  	Y = CLng(Right(XY, Len(XY) - iZB))
					  	//MsgBox x &"and"&y
					    	If X > 0 And Y > 0 Then 
						  	   Exit For
						    End If
						Delay 1000 * 10//等待10秒
						ux1=ux1 + 5 
						uy1=uy1 + 5 	
					  	Next	
		    		  
		    		  End If  
				  Else 
		    		//MsgBox "程序错误,查找不到轨迹记录！进入测试页"
		    	  End If
						
			Else 
			  Exit For //查找不到讲课则退出到下一章
			End If
		Next
		
		
	Else 
	  Exit For
	End If	
Next


Sub FINDF4()
	MoveTo ux + 300, uy + 220
	LeftDown 1
	Delay 250
	LeftUp 1
	//点击页面空白
	KeyDown 17, 1
	KeyPress 70, 1
	KeyUp 17, 1
	Delay 500
	//ALT+F查找关键字
    KeyPress "Delete", 7
    //删除原有关键字
End Sub

Sub nextpage()
	MoveTo ux + 400, uy + 220
	LeftDown 1
	Delay 250
	LeftUp 1
	MoveTo ux+100, uy+500
	Delay 250
For k = 0 To 20
	MouseWheel - 50 
	Delay 250
	k=k+1
Next
End Sub

Sub lastpage()
	MoveTo ux + 400, uy + 220
	LeftDown 1
	Delay 250
	LeftUp 1
	MoveTo ux+100, uy+500
	Delay 250
For k = 0 To 20
	MouseWheel  50 
	Delay 250
	k=k+1
Next
End Sub




