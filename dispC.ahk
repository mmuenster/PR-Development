#SingleInstance, Force
#Persistent

SetTimer, CheckCaseCount, 1800000
Gosub, CheckCaseCount
return


^!5::
{
	distributedCount=141
	totalJarsInQueue=411
	FormatTime, TimeString
	URL      := "https://dazzling-torch-3393.firebaseio.com/newTest.json"
	POSTData="%distributedCount% of %totalJarsInQueue% distributed -- %TimeString%"
	Msgbox, %POSTData%

	j := URLPut(URL, PostData)

	return
}

CheckCaseCount:
{

URLDownloadToFile, http://s-irv-autoasgn/autoassign2/report_path_case_status.php, distHtml.txt

FileRead, html, distHTML.txt
FileDelete, distHTML.txt

;  Login info  Username=mmuenster  Password=mmuenster@2015

document := ComObjCreate("HTMLfile")
document.write(html)
all := document.getElementsByTagName("table")
	
Sleep, 1000
tempVar := 2
table := all[tempVar]

notDistributedCount := distributedCount := 0

Loop, % table.rows.length - 1
{	
	tempVar := A_Index-1
	if (table.rows[tempVar].cells[1].innerHTML>0 AND table.rows[tempVar].cells[2].innerHTML<>"&nbsp;")
		distributedCount += table.rows[tempVar].cells[1].innerHTML
	else
		notDistributedCount += table.rows[tempVar].cells[1].innerHTML
	
}

	totalJarsInQueue := distributedCount + notDistributedCount
	;Msgbox, Distribution Summary`n-------------------------`n %distributedCount% of %totalJarsInQueue% distributed

FormatTime, TimeString
URL      := "https://dazzling-torch-3393.firebaseio.com/test.json"
POSTData="%distributedCount% of %totalJarsInQueue% distributed -- %TimeString%"
;Msgbox, %POSTData%

j := URLPut(URL, PostData)

return
}

UrlPut(URL, data) 
{
   WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
   WebRequest.Open("Put", URL, false)
   WebRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
   WebRequest.Send(data)
   Return WebRequest.ResponseText
}
