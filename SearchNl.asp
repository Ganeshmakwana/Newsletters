<%
cSearchStr = Request("SearchStr")
'
ConnStr = "Driver=Microsoft Visual Foxpro Driver;UID=;SourceType=DBC;SourceDB=D:\AP\NewsLetter\nldbFox.dbc"
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open ConnStr
SqlStr = "SELECT * From Letters Order By PostDate Where nlsend = 'Y' AND '"+cSearchStr+"' $ Search"
Set RS = Conn.Execute(SqlStr)
'
If Rs.Eof OR Rs.Bof Then
 errormessage = "No Topics available for Referance, Search Text : "+cSearchStr
Else
 errormessage = "Following Topics Found, Search Text : "+cSearchStr
End if
%>
<html>

<head>
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

</td></tr><!--msnavigation--></table><!--msnavigation--><table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><!--msnavigation--><td valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber1">
        <tr>
          <td width="17%">&nbsp;</td>
          <td width="83%" align="right"><a href="../default.htm">KMEPL Home Page</a></td>
        </tr>
        <tr>
          <td width="17%">&nbsp;</td>
          <td width="83%" align="right"><a href="Default.asp">News Letters Home Page</a></td>
        </tr>
        <tr>
          <td width="17%">&nbsp;</td>
          <td width="83%"><H1><%=errormessage%></H1></td>
        </tr>
      </Table><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber2">
        <%If Rs.Eof OR Rs.Bof Then%>
        <tr>
          <td width="17%">&nbsp;</td>
          <td width="83%" style="border-bottom-style: solid; border-bottom-width: 1"><h2><%=errormessage%></h2> </td>
        </tr>
        <%End if%>
        <%Do While Not Rs.Eof%>
        <tr>
          <td width="17%" style="border-right-style: solid; border-right-width: 1">&nbsp;</td>
          <td width="83%" ColSpan=2 bgcolor="#CCFFFF" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1"><h2><a target="_blank" href="readit.asp?cID=<%Response.write cStr(Rs.Fields("ID"))%>"><%Response.write Ltrim(Rtrim(Rs.Fields("Title")))%></a></h2>    <Font Size=1><%Response.write FormatDateTime(Rs.Fields("PostDate"),2)+"-"+Rs.Fields("PostTime")%></Font></a></td>
      </tr>
      <tr>
        <td width="17%" style="border-right-style: solid; border-right-width: 1">&nbsp;</td>
        <td width="83%" ColSpan=2 style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1">Author : <Font Size=2><%Response.write Ltrim(Rtrim(Rs.Fields("Author")))%></Font></td>
      </tr>
      <tr>
        <td width="17%" style="border-right-style: solid; border-right-width: 1">&nbsp;</td>
        <td width="83%" ColSpan=2 style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1">Category : <Font Size=2><%Response.write Ltrim(Rtrim(Rs.Fields("Category")))%></Font></td>
      </tr>
      <tr>
        <td width="17%" style="border-right-style: solid; border-right-width: 1">&nbsp;</td>
        <td width="83%" ColSpan=2 style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1">Number of Hits : <Font Size=2><%Response.write Rs.Fields("Hits")%></Font></td>
      </tr>
      <tr>
        <td width="17%" style="border-right-style: solid; border-right-width: 1">&nbsp;</td>
        <td width="42%" align="left" style="border-left-style: solid; border-left-width: 1; border-bottom-style: solid; border-bottom-width: 1">Read Comment(0)</td>
        <td width="41%" align="left" style="border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1; border-bottom-style: solid; border-bottom-width: 1">write Comment(0)</td>
      </tr>
      <tr>
        <td width="17%">&nbsp;</td>
        <td width="83%"  ColSpan=2 align="right">-</td>
      </tr>
      <%
Rs.MoveNext
Loop
%>
    </table><!--msnavigation--></td></tr><!--msnavigation--></table><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

</td></tr><!--msnavigation--></table></body>
<%Rs.Close%>
</html>