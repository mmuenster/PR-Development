#SingleInstance Force
ComObjError(false)
Menu, Tray, Tip, KeepActive

Gosub, ^!8
SetTimer, KeepActive, 1000000

return


KeepActive:
{
 
	MouseGetPos, x, y
	MouseMove, x +30, y+30
	MouseMove, x, y
	
   IfWinNotExist, TeamViewer
	{
		Gosub, ^!8
	}
   return
}

F6::
{
Process,Close,WinSURGE.exe
Process,WaitClose,WinSURGE.exe,2
If(!ErrorLevel)
{
	Run, "C:\Program Files (x86)\WinSURGE\WinSURGE.exe"
	WinWaitActive, WinSURGE
	Send, Spr{!}ng2018
	Send, {Enter}
	WinWaitActive, Login Message
	Send, {Enter}
	
	WinWaitActive, WinSURGE, Pathologist-CR
	Click, 760, 100
}
else
{
	Msgbox, Could not close WinSurge
}
return
}

^!j::
{
	totalJarCount := 0
	daysInMonths := [31,29,31,30,31,30,31,31,30,31,30,31 ]
	InputBox, monthToSearch, What Month and Year Do You Want To Search? MM-YYYY
	StringSplit, m, monthToSearch, -
	monthSearch := m1
	yearSearch := m2
	StringLeft, y, A_UserName, 5
	StringUpper, y, y
	ComObjError(True)
	s = Select p.name, p.id, p.state, p.suid, p.wsid from pathologists p where p.abbr='%y%'
	CodeDataBaseQuery(s)
	pathSearch := Result_5
	numdays := daysInMonths[monthSearch]
	Loop, %numdays%
	{
		jarcount:=casecount:=mxcount:=totalSlideCount:=0

		if (A_Index<10)
			daySearch=0%A_Index%
		else
			daysearch=%A_Index%
		
		FormatTime, dow, %yearSearch%%monthSearch%%daySearch%, dddd
		todaydate = %yearSearch%-%monthSearch%-%daySearch%
		s := "select s.number, s.numberofspecimenparts, s.dx, s.calculatedslidecount from specimen s where s.path =" . pathSearch . " and s.sodate = '" . todaydate . "'"
		WinSurgeQuery(s)
		Loop, parse, msg, `n 
		{
			If(A_LoopField)
			{
				
				StringSplit, res, A_LoopField, ¥
				casecount := casecount + 1
				jarcount := jarcount + res3
				totalSlideCount := totalSlideCount + res5
				lowerCaseProblem := RegExMatch(res4, "%%P%% [a-z]")
				if lowerCaseProblem>0
					{
					SoundBeep
					Msgbox,  %res2%`nSTOP! STOP! STOP!
				}

				IfInString, res2, MX
					mxcount := mxcount + 1
			}
		}
		
		daysJarCount := jarcount + mxcount*2
		if(dow="Saturday" OR dow="Sunday")
			totalJarCount := totalJarCount + daysJarCount
		else if (daysJarCount>142)
			totalJarCount := totalJarCount + daysJarCount - 142
		
		appendText := todaydate . "," . dow . "," . daysJarCount . "," . totalJarCount . "," . totalSlideCount . "`n"
	
	FileAppend, %appendText%, C:\Users\mmuenster\Documents\jarcountsummary.csv 
	}

	Msgbox, Done!
	return	
}

^!8::
{
	IfWinNotExist, TeamViewer
	{
		Run, C:\Users\mmuenster\Documents\TeamViewer\TeamViewer.exe
		WinWait, TeamViewer, , 5
		WinActivate, TeamViewer
		WinWaitActive, TeamViewer
		Loop, 
		{
			sleep, 500
			ControlGetText, tempVar,Edit2 , TeamViewer
			tempVar1 := RegExMatch(tempVar, "\d+")
			If tempVar1
				break
		}
		MouseClick, left, 243, 308
		Sleep, 500
		MouseClick, left, 306, 385
		WinWaitActive, TeamViewer options
		ControlSetText, Edit10, 1008, TeamViewer options
		ControlSetText, Edit11, 1008, TeamViewer options
		MouseClick, left, 669,558
		WinWaitClose, TeamViewer options
		WinActivate, TeamViewer
		WinMinimize, Teamviewer
	}

	
return
}

^!7::
{
InputBox, s, Enter the SQL phrase you want...
clipboard = %s%
WinSurgeQuery(s)
If msg
	Msgbox, Results:`n%msg%
return
}

WinSurgeQuery(s)
{
	global
	Loop, 15          ;Blank the results array
		Result_%A_Index% =

	connectstring := "DRIVER={InterSystems ODBC};SERVER=dfw-mpwsg01;DATABASE=wins1csql;uid=sqluser;pwd=sqluserpw;"
	adodb := ComObjCreate("ADODB.Connection")
	rs := ComObjCreate("ADODB.Recordset")
	rs.CursorType := "0"
	adodb.open(connectstring)
RetryWinSURGEDatabase:
	rs := adodb.Execute(s)
	If A_LastError
		{
			Msgbox, 4101, Error Message, There was an error accessing the WinSURGE Database.`n`ns=%s%`n  A_LastError = %A_LastError%
			IfMsgbox, Retry
				Goto, RetryWinSURGEDatabase
			else ifMsgBox, Cancel
				return
		}	
	msg := ""
	txt := rs.state
	If !txt
		return
	while rs.EOF = 0{
		for field in rs.fields
			msg := msg . "¥" . Field.Value
		msg = %msg%`n
		rs.MoveNext()
	}
	
	Loop, parse, msg, ¥ 
		{
		if (A_Index = 1)
			Continue
		Else
			{
			 i := A_Index -1 
			 Result_%i% := A_LoopField		
			}
		}
	
	rs.close()   
	adodb.close()
	return 	
}

CodeDatabaseQuery(s)  ;((((((((((((((((
{
global
RetryCodeDatabase:
	Loop, 15          ;Blank the results array
		Result_%A_Index% =
		
	connectstring := "DRIVER={SQL Server};SERVER=s-irv-sql02;DATABASE=winsurgehotkeys;uid=wshotkeys;pwd=hotkeys10;"
	adodb := ComObjCreate("ADODB.Connection")
	adodb.open(connectstring)

	rs := adodb.Execute(s)
	If A_LastError
		{
			Msgbox, 4101, Error Message, There was an error accessing the Code Database.`n  `ns=%s%`n  A_LastError = %A_LastError%
			IfMsgbox, Retry
				Goto, RetryCodeDatabase	
			else ifMsgBox, Cancel
				ExitApp
		}	
		
	msg := ""
	txt := rs.state
	If !txt
		return
		
	while rs.EOF = 0{
		for field in rs.fields
			msg := msg . "¥" . Field.Value
		rs.MoveNext()
	}
	
	Loop, parse, msg, ¥ 
		{
		if (A_Index = 1)
			Continue
		Else
			{
			 i := A_Index -1 
			 Result_%i% := A_LoopField		
			}
		}

	rs.close()   
	adodb.close()
	return msg	
}


^!r::Reload