	s = Update pathologists set ttwentyfive03='MRM' where abbr='rlash'
	CodeDataBaseQuery(s)

	Msgbox, msg=%msg%

CodeDatabaseQuery(s)  
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

