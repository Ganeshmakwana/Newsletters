<%
netUser = " "
netUser = Request.ServerVariables("LOGON_USER")
netUserIp = cSTR(Request.ServerVariables("REMOTE_ADDR"))
StartFM = InStr(1,netUser,"\",1)+1
netUser = MID(netUser, StartFM ,(Len(Ltrim(Rtrim(netUser)))- StartFM)+1)
'
ConnStrX = "Driver=Microsoft Visual Foxpro Driver;UID=;SourceType=DBC;SourceDB=D:\AP\NewsLetter\nldbFox.dbc"
Set ConnX = Server.CreateObject("ADODB.Connection")
ConnX.Open ConnStrX
SqlStrX = "SELECT max(id) AS MaxID From Letters Order By Letters.Id"
Set RSX = ConnX.Execute(SqlStrX)
'
If Not Rsx.Eof Then
 If cCur(Rsx("Maxid")) > 0 Then
  nMaxId = cCur(Rsx("Maxid"))+1
 Else
  nMaxId = 1
 End if
Else
 nMaxid = 1
End if
RsX.Close
ConnX.Close
Response.write nMaxid
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="Microsoft Theme" content="none, default">
<meta name="Microsoft Border" content="tb">
</head>
<body><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<p align="center"><font size="6"><strong>
</strong></font><br>
</p>
<p align="center">&nbsp;</p>

</td></tr><!--msnavigation--></table><!--msnavigation--><table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><!--msnavigation--><td valign="top">



      <%
Dim Upload
cSuccess = ""
Set Upload = Server.CreateObject("TABS.Upload")
Upload.Start "C:\TEMP"
Upload.MaxBytesToAbort = 10 * 1024 * 1024
If Upload.Form("UpLoadFile").FileSize <> 0 Then
    Set Conn = Server.CreateObject("ADODB.connection")
    ConnStr= "Driver=Microsoft Visual Foxpro Driver;UID=;SourceType=DBC;SourceDB=D:\AP\NewsLetter\nldbFox.dbc"
    Set cmdTemp = Server.CreateObject("ADODB.Command")
    Conn.Open ConnStr
  	Set Rs = Server.CreateObject("ADODB.Recordset")
      cmdTemp.CommandText = "Letters"
      cmdTemp.CommandType = 2
     Set cmdTemp.ActiveConnection = Conn
 	 Rs.Open cmdTemp, , 2,3
     Rs.AddNew
	 Rs("Id") = nMaxId
	 Rs("DocName") = Upload.Form("uploadFile").FileName
	 Upload.Form("uploadFile").SaveAsBlob Rs("DocData")
	 Rs("Title") = Rtrim(Ltrim(Upload.Form("Title")))
	 Rs("Author") = Rtrim(Ltrim(Upload.Form("Author")))
	 Rs("Tpclass") = Rtrim(Ltrim(Upload.Form("Tclass")))
	 Rs("Category") = Rtrim(Ltrim(Upload.Form("Category")))
	 Rs("PostDate") = FormatDateTime(Date,2)
	 Rs("PostTime") = cStr(Hour(Now))+":"+cStr(Minute(Now))+":"+cStr(Second(Now))
	 Rs("NetUserNm") = netUser
	 Rs("NetUserIP") = NetUserIP
	 Rs("Search")	= Upload.Form("SearchTxt")
	 Rs("Hits") = 0
	 Rs("Nlsend") = "Y"
	 Rs.Update
	 Rs.Close
	cSuccess = "Upload Complete"
Else
	cSuccess = "No file exists"
End If
%><table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%" align="center">
            <p align="right"><a href="../default.htm">KMEPL Home Page</a></td>
          </tr>
          <tr>
            <td width="100%" align="center">
              <p align="right"><a href="Default.asp">News Letters Home Page</a></td>
            </tr>
            <tr>
              <td width="100%" align="center">&nbsp;</td>
            </tr>
            <tr>
              <td width="100%" align="center">Topic <%=Rtrim(Ltrim(Upload.Form("Title")))%></td>
            </tr>
            <tr>
              <td width="100%" align="center">Successfully uploaded in Database.</td>
            </tr>
            <tr>
              <td width="100%" align="center"><a href="Default.htm">click here to go Back to List of Topics</a></td>
            </tr>
          </table>
          <%
Set Upload = Nothing
Conn.Close
%> <!--msnavigation--></td></tr><!--msnavigation--></table><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

</td></tr><!--msnavigation--></table></body>
</html>