<%
netUser = " "
netUser = Request.ServerVariables("LOGON_USER")
netUserIp = cSTR(Request.ServerVariables("REMOTE_ADDR"))
StartFM = InStr(1,netUser,"\",1)+1
netUser = MID(netUser, StartFM ,(Len(Ltrim(Rtrim(netUser)))- StartFM)+1)
'
cid = Request("cId")
'
ConnStr = "Driver=Microsoft Visual Foxpro Driver;UID=;SourceType=DBC;SourceDB=D:\AP\NewsLetter\nldbFox.dbc"
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open ConnStr
SqlStr = "SELECT * From Letters Where Id = "+ cStr(cId)
Set RS = Conn.Execute(SqlStr)
'
cDocName = Rs.Fields("DocName")
tpclass = Rs.Fields("tpclass")
category = Rs.Fields("category")
'
Set Conn1 = Server.CreateObject("ADODB.Connection")
Conn1.Open ConnStr
SqlStr1 = "Update Letters Set Hits = Hits+1  Where Id = "+ cStr(cid)
conn1.execute SqlStr1
Conn1.Close
'
Set Conn2 = Server.CreateObject("ADODB.Connection")
Conn2.Open ConnStr
SqlStr2 = "Insert Into thitlist (DocName, NetUserNm, NetUserip, tpclass, Category, readdate, readtime) Values('"+cDocName+"','"+NetUser+"','"+NetUserIp+"','"+tpclass+"','"+category+"',Date("+Mid(FormatDateTime(Date,2),7,4)+","+Mid(FormatDateTime(Date,2),4,2)+","+Mid(FormatDateTime(Date,2),1,2)+"),'"+cStr(Hour(Now))+":"+cStr(Minute(Now))+":"+cStr(Second(Now))+"')"
conn2.execute SqlStr2
Conn2.Close
If Not Rs.EOF Then
	Set Download = Server.CreateObject("TABS.Download")
	Download.FileName = Rs("DocName")
	'Transfer data to the web browser.
	Download.TransferBlob Rs("DocData"), False
End If
Rs.Close
Conn.Close
%>