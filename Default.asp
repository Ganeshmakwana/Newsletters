<%
csCategory = Request("sCategory")
csTpClass = Request("sTpclass")
'*if isEmpty(csCategory) Then
'* csCategory = "All"
'*End if
'*if isEmpty(csTpClass) Then
'* csTpClass = "General"
'*End if
'
ConnStr = "Driver=Microsoft Visual Foxpro Driver;UID=;SourceType=DBC;SourceDB=D:\AP\NewsLetter\nldbFox.dbc"
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open ConnStr
If csCategory = "All" Then
 SqlStr = "SELECT * From Letters Where Letters.nlsend = 'Y' Order By Letters.PostDate"
Else
 SqlStr = "SELECT * From Letters Where (Letters.nlsend = 'Y') AND (Letters.Category = '"+ csCategory + "') Order By Letters.PostDate"
End if
'
Set RS = Conn.Execute(SqlStr)
'
If Rs.Eof OR Rs.Bof Then
 errormessage = "No Letters available for Referance"
End if
'
Set Conn1 = Server.CreateObject("ADODB.Connection")
Conn1.Open ConnStr
'
If csTpclass = "All" Then
 SqlStr1 = "SELECT * From Cats Order By Category"
Else
 SqlStr1 = "SELECT * From Cats Where TpClass = '"+csTpclass+"' Order By Category"
End if
'
Set RS1 = Conn1.Execute(SqlStr1)
'
Set Conn2 = Server.CreateObject("ADODB.Connection")
Conn2.Open ConnStr
'
SqlStr2 = "SELECT * From TClass Order By Tpclass"
'
Set RS2 = Conn2.Execute(SqlStr2)
'
cColSpan=0
Do While Not Rs2.eof
 cColSpan = cColSpan +1
Rs2.MoveNext
Loop
Rs2.MoveFirst
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="Microsoft Theme" content="none, default">
<meta name="Microsoft Border" content="none, default">
</head>

<body>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber3">
        <tr>
          <td align="right" colspan="<%=cColSpan%>"><a href="../default.htm">KMEPL Home Page</a></td>
        </tr>
        <tr>
	      <%Do While Not Rs2.Eof%>
          <td align="Center"><a href="Default.asp?sTpclass=<%=Rs2.Fields("TpClass")%>"><%=Rs2.Fields("TpClass")%></a>&nbsp;</td>
          <%Rs2.MoveNext
    Loop%>
        </tr>
        <tr>
          <td align=Center colspan="<%=cColSpan%>">
            <form method="POST" action="Default.asp">
              <select size="1" name="sCategory" tabindex="8">
	            <%Do While Not Rs1.Eof%>
                <option value="<%Response.write Rs1.Fields("Category")%>"><%Response.write Rs1.Fields("Category")%></option>
	            <%Rs1.MoveNext
	Loop%>
              </select><input type="submit" value="Submit" name="B1" tabindex="9">
            </form>
            </td>
        </tr>
      </table>
      <p></p>
      <table border="2" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber1">
        <tr>
          <td colspan="3"><h2><%=csCategory%> - Recent Topics</h2> </td>
        </tr>
        <%If Rs.Eof OR Rs.Bof Then%>
        <tr>
          <td colspan="3"><h2><%=errormessage%></h2> </td>
        </tr>
        <%End if%>
<tr>
<td>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber2">
<%Do While Not Rs.Eof%>
     <tr>
       <td width="100%" ColSpan=2 bgcolor="#CCFFFF"><h2><a target="_blank" href="readit.asp?cId=<%Response.write cStr(Rs("Id"))%>"><%Response.write Ltrim(Rtrim(Rs.Fields("Title")))%></a></h2><Font Size=1><%Response.write FormatDateTime(Rs.Fields("PostDate"),2)+"-"+Rs.Fields("PostTime")%></Font></a></td>
     </tr>
     <tr>
       <td width="100%" ColSpan=2>Author : <Font Size=2><%Response.write Ltrim(Rtrim(Rs.Fields("Author")))%></Font></td>
     </tr>
     <tr>
       <td width="100%" ColSpan=2>Category : <Font Size=2><%Response.write Ltrim(Rtrim(Rs.Fields("Category")))%></Font></td>
     </tr>
     <tr>
       <td width="100%" ColSpan=2>Number of Hits : <Font Size=2><%Response.write Rs.Fields("Hits")%></Font></td>
     </tr>
     <tr>
       <td width="100%"  ColSpan=2 align="left">Read Comment(0) write Comment(0)</td>
     </tr>
     <tr>
       <td width="100%"  ColSpan=2 align="right">-</td>
     </tr>
<%
Rs.MoveNext
Loop
%>
</table>
</td>
</tr>
</table>
</body>
<%
Rs.Close
Conn.Close
Rs1.Close
Conn1.Close
Rs2.Close
Conn2.Close
%>