<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/guest.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_guest_STRING
    MM_editCmd.CommandText = "DELETE FROM guest WHERE ID = ?"
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, Request.Form("MM_recordId")) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "admin.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If

End If
%>
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="login.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_guest_STRING
Recordset1_cmd.CommandText = "SELECT * FROM guest WHERE ID = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>删除留言</title>
<link href="style.css" rel="stylesheet" type="text/css" />
<style type="text/css">
img {
	height: 50px;
	width: 50px;
}
</style>
<link href="../SpryAssets/SpryValidationTextarea.css" rel="stylesheet" type="text/css" />
<link href="../SpryAssets/SpryValidationTextField.css" rel="stylesheet" type="text/css" />
<script type="text/javascript">
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
</script>
</head>


<body>

  <table width="640"align="center" bgcolor="#CCCCCC"  cellpadding="5" >
    <tr class="beijing">
      <td class="biaoti" align="center">删除留言</td>
    </tr>
    <tr>
      <td align="right"><a href="index.asp">
      
      <p>&nbsp;</p>
      <img src="image/view.jpg"  /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="insert.asp"><img src="image/add.jpg"  /></a></td>
    </tr>
    <tr>
     <td align="center"><p>你确定要删除留言吗？</p>
       
       <form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
         <input type ="submit" name ="submit" value ="确定删除"/>
&nbsp;&nbsp;&nbsp;
<input name ="reset" type ="reset" onclick="MM_goToURL('parent','admin.asp');return document.MM_returnValue" value ="取消删除"  />
<input type="hidden" name="MM_delete" value="form1" />
<input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("ID").Value %>" />
       </form>
      </td>
    </tr>
      
    <tr>
      <td align="center">
      <input name="Date" type="hidden" id="Date" /></td>
</tr>
    <tr>
      <td class="zhengwen"  align="center">版权所有17教育技术学</td>
    </tr>
  </table>
<script type="text/javascript">
var sprytextarea1 = new Spry.Widget.ValidationTextarea("sprytextarea1");
var sprytextfield1 = new Spry.Widget.ValidationTextField("sprytextfield1");
var sprytextfield2 = new Spry.Widget.ValidationTextField("sprytextfield2");
var sprytextfield3 = new Spry.Widget.ValidationTextField("sprytextfield3", "email");
var sprytextfield4 = new Spry.Widget.ValidationTextField("sprytextfield4", "integer");
</script>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
