html     := ""
URL      := "https://dazzling-torch-3393.firebaseio.com/test.json"
POSTData := "46"

j := URLPost(URL, PostData)

Msgbox, %j%
j := URLPut(URL, PostData)

Msgbox, %j%


UrlPost(URL, data) {
   WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
   WebRequest.Open("POST", URL, false)
   WebRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
   WebRequest.Send(data)
   Return WebRequest.ResponseText
}

UrlPut(URL, data) {
   WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
   WebRequest.Open("Put", URL, false)
   WebRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
   WebRequest.Send(data)
   Return WebRequest.ResponseText
}
