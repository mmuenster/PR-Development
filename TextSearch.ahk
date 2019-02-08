
    InputBox, searchdate, Search Date "MM-DD", Please enter the date to search.  Leave blank for today!, , 640, 480
    InputBox, textToSearch, Diagnosis Search Text, Please enter the text you wish to search for. , , 640, 480
	
	if (searchdate)
		todaydate = %A_YYYY%-%searchdate%
	else
		todaydate = %A_YYYY%-%A_MM%-%A_DD%
	
	s := "select s.number, s.numberofspecimenparts, s.dx from specimen s where s.dx LIKE '%" . textToSearch . "%' and s.sodate >= '" . todaydate . "'"
	;Msgbox, %s%
	
	WinSurgeQuery(s)

	Msgbox, %msg%
	return
	
	
	WinSurgeQuery(s)
{
	global
	Loop, 15          ;Blank the results array
		Result_%A_Index% =

	;connectstring := "DRIVER={InterSystems ODBC};SERVER=s-irv-wsg01;DATABASE=wins1csql;uid=_system;pwd=sys;"
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
