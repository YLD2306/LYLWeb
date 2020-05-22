<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/guest.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("name"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "admin.asp"
  MM_redirectLoginFailed = "fail.asp"

  MM_loginSQL = "SELECT [user], pass"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM [admin] WHERE [user] = ? AND pass = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_guest_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 15, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 32, Request.Form("password")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>留言本管理登录</title>
<link href="style.css" rel="stylesheet" type="text/css" />
<style type="text/css">
img {
	height: 50px;
	width: 50px;
}
</style>
<link href="../SpryAssets/SpryValidationTextarea.css" rel="stylesheet" type="text/css" />
<link href="../SpryAssets/SpryValidationTextField.css" rel="stylesheet" type="text/css" />
</head>


<body>

  <table width="640"align="center"  bgcolor="#CCCCCC"  cellpadding="5">
    <tr class="beijing">
      <td class="biaoti" align="center">留言本管理登录</td>
    </tr>
    <tr>
      <td align="right"><a href="index.asp"><img src="image/view.jpg"  /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="insert.asp"><img src="image/add.jpg"  /></a></td>
    </tr>
    <tr>
      <td><form ACTION="<%=MM_LoginAction%>" METHOD="POST" id="form1" name="form1">
       <table width="100%" border="0">
  <tr>
    <td width="41%" align="right">账号：</td>
    <td width="59%"><label for="name"></label>
      <input name="name" type="text" id="name" size="20" /></td>
  </tr>
  <tr>
    <td align="right">密码：</td>
    <td><label for="password"></label>
      <input name="password" type="password" id="password" size="20" /></td>
  </tr>
  <tr>
    <td align="right">&nbsp;</td>
     <td ><input type ="submit" name ="submit" value ="提交"/>       
        &nbsp;&nbsp;&nbsp;
        <input type ="reset" name ="reset" value ="重置"  /></td>
    
    </tr>
</table>

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
