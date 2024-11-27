<%
no = 0
subject = "Test 1"
message = "Hi Ganesh"
message = message & vbcrlf & vbcrlf
message = message & "To stop receiving emails click here :" 
message = message & vbcrlf
message = message & "http://yoursite.com/urfolder/del.asp?email="
Set mail = Server.CreateObject("CDONTS.NewMail")
mail.From = "rao@kmohan.com"
mail.To = "ganesh@kmohan.com"
mail.Subject = subject
mail.Body = message
mail.Send
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="Microsoft Theme" content="none, default">
<meta name="Microsoft Border" content="none, default">
</head>

<body>
&nbsp;</body></html>