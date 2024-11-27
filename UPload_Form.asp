<%
netUser = " "
netUser = Request.ServerVariables("LOGON_USER")
StartFM = InStr(1,netUser,"\",1)+1
netUser = MID(netUser, StartFM ,(Len(Ltrim(Rtrim(netUser)))- StartFM)+1)
Netuser = Ucase(NetUser)
'
ConnStr = "Driver=Microsoft Visual Foxpro Driver;UID=;SourceType=DBC;SourceDB=D:\AP\NewsLetter\nldbFox.dbc"
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open ConnStr
SqlStr = "SELECT Category From Cats Order By Category"
'
Set RS = Conn.Execute(SqlStr)
'
Set Conn1 = Server.CreateObject("ADODB.Connection")
Conn1.Open ConnStr
SqlStr1 = "SELECT Tpclass From Tclass Order By Tpclass"
'
Set RS1 = Conn1.Execute(SqlStr1)
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

</td></tr><!--msnavigation--></table><!--msnavigation--><table dir="ltr" border="0" cellpadding="0" cellspacing="0" width="100%"><tr><!--msnavigation--><td valign="top"><!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.Title.value == "")
  {
    alert("Please enter a value for the \"Topic Title\" field.");
    theForm.Title.focus();
    return (false);
  }

  if (theForm.Title.value.length < 2)
  {
    alert("Please enter at least 2 characters in the \"Topic Title\" field.");
    theForm.Title.focus();
    return (false);
  }

  if (theForm.Title.value.length > 100)
  {
    alert("Please enter at most 100 characters in the \"Topic Title\" field.");
    theForm.Title.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" name="FrontPage_Form1" enctype="multipart/form-data" action="UPload.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber1">
        <tr>
          <td width="100%" align="center" colspan="3">
            <p align="right"><a href="../default.htm">KMEPL Home Page</a></td>
          </tr>
          <tr>
            <td width="100%" align="center" colspan="3">
              <p align="right"><a href="Default.asp">News Letters Home Page</a></td>
            </tr>
            <tr>
              <td width="31%" align="center">&nbsp;</td>
              <td width="69%" colspan="2">&nbsp;</td>
            </tr>
            <tr>
              <td width="31%" align="center">Title</td>
              <td width="69%" colspan="2">
              <!--webbot bot="Validation" s-display-name="Topic Title" b-value-required="TRUE" i-minimum-length="2" i-maximum-length="100" --><input type="text" name="Title" size="65" tabindex="1" maxlength="100"></td>
            </tr>
            <tr>
              <td width="31%" align="center">Author</td>
              <td width="69%" colspan="2"><input type="text" name="Author" size="50" tabindex="2"></td>
            </tr>
            <tr>
              <td width="100%" align="center" colspan="3">&nbsp;</td>
            </tr>
            <tr>
              <td width="31%" align="center">Class</td>
              <td width="69%" colspan="2"><select size="1" name="Tclass" tabindex="3">
	              <%Do While Not Rs1.EOf %>
                  <option value="<%Response.write Rs1.Fields("Tpclass")%>"><%Response.write Rs1.Fields("Tpclass")%></option>
                  <%Rs1.MoveNext
    Loop%></select></td>
            </tr>
            <tr>
              <td width="31%" align="center">Category</td>
              <td width="69%" colspan="2"><select size="1" name="Category" tabindex="4">
	              <%Do While Not Rs.EOf %>
                  <option value="<%Response.write Rs.Fields("Category")%>"><%Response.write Rs.Fields("Category")%></option>
                  <%Rs.MoveNext
    Loop%></select></td>
            </tr>
            <tr>
              <td width="100%" align="center" colspan="3">&nbsp;</td>
            </tr>
            <tr>
              <td width="31%" align="center">File Name</td>
              <td width="69%" colspan="2"><input type="file" name="UpLoadFile" size="30" tabindex="5"></td>
            </tr>
            <tr>
              <td width="100%" align="center" colspan="3">&nbsp;</td>
            </tr>
            <tr>
              <td width="31%" align="center">Search String</td>
              <td width="69%" colspan="2"><textarea rows="5" name="SearchTxt" cols="55" tabindex="6"></textarea></td>
            </tr>
            <tr>
              <td width="100%" colspan="3">&nbsp;</td>
            </tr>
            <tr>
              <td width="31%" align="center">&nbsp;</td>
              <td width="35%" align="center"><input type="submit" value="Upload" Name="B1" tabindex="7" style="float: left"></td>
              <td width="34%" align="center"><input type="reset" value="Reset" name="B2" tabindex="8"></td>
            </tr>
          </table>
          </form>

          
    
    
  <!--msnavigation--></td></tr><!--msnavigation--></table><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

</td></tr><!--msnavigation--></table></body>
  <%
Rs.Close
Rs1.Close
Conn.Close
Conn1.Close
%>
</html>