' copyright 2011 cenan ozen <cenan.ozen@gmail.com>
' see README for explanation

On Error Resume Next

Const smtpServer = ""
Const smtpServerPort = 25
Const senderUserName = ""
Const senderPassword = ""

Function pingTarget(strTarget)
	Dim strPingResults
	Set WshShell = WScript.CreateObject("WScript.Shell")
	PINGFlag = Not CBool(WshShell.run("ping -n 3 -w 1000 " & strTarget,0,True))
	If PINGFlag = True Then
		pingTarget = TRUE   
	Else
		pingTarget = FALSE
	End If
End Function

Sub sendMail(mailBody, mailTo)
	Set objMessage = CreateObject("CDO.Message")
	objMessage.Subject = "ping report"
	objMessage.From = senderUserName
	objMessage.To = mailTo
	objMessage.TextBody = mailBody
	objMessage.Configuration.Fields.Item _
	   ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objMessage.Configuration.Fields.Item _
	   ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = smtpServerPort
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	objMessage.Configuration.Fields.Item _
	   ("http://schemas.microsoft.com/cdo/configuration/sendusername") = senderUserName
	objMessage.Configuration.Fields.Item _
	   ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = senderPassword
	objMessage.Configuration.Fields.Update
	objMessage.Send
End Sub

Sub Main()
	Dim fso, fin, i, res
	Dim iplist()
	i = 0
	Set fso = CreateObject("Scripting.FileSystemObject")
	set fin = fso.OpenTextFile("iplist.txt", 1, false)
	Do Until fin.AtEndOfStream
		ReDim Preserve iplist(i)
		iplist(i) = fin.ReadLine
		i = i + 1
	Loop
	fin.Close
	res = ""
	For l = LBound(iplist) to UBound(iplist)
		If pingTarget(iplist(l)) = False Then
			res = res & iplist(l) & " does not respond to ping" & vbcrlf
		End If
	Next
	If res <> "" Then
		sendMail res, "you@yourmailaddress.com"
		' add more emails here
	End If
End Sub
