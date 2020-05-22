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
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_insert")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the insert
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_guest_STRING
    MM_editCmd.CommandText = "INSERT INTO guest (Name, Email, Homepage, QQ, fromwhere, ICON, Content, [Date]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 202, 1, 15, Request.Form("Name")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 32, Request.Form("Email")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 30, Request.Form("Homepage")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("QQ")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 202, 1, 32, Request.Form("fromwhere")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 202, 1, 60, Request.Form("ICON")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param7", 203, 1, 1073741823, Request.Form("Content")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param8", 135, 1, -1, MM_IIF(Request.Form("Date"), Request.Form("Date"), null)) ' adDBTimeStamp
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "index.asp"
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
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>签写留言</title>
<link href="style.css" rel="stylesheet" type="text/css" />
<script src="../SpryAssets/SpryValidationTextarea.js" type="text/javascript"></script>
<script src="../SpryAssets/SpryValidationTextField.js" type="text/javascript"></script>
<script>  
	//自动关闭提示框
	function Alert(str) {    
		var msgw,msgh,bordercolor;    
		msgw=350;//提示窗口的宽度    
		msgh=80;//提示窗口的高度    
		titleheight=25 //提示窗口标题高度    
		bordercolor="#336699";//提示窗口的边框颜色    
		titlecolor="#99CCFF";//提示窗口的标题颜色    
		var sWidth,sHeight;    
		//获取当前窗口尺寸    
		sWidth = document.body.offsetWidth;    
		sHeight = document.body.offsetHeight;    
	//    //背景div    
		var bgObj=document.createElement("div");    
		bgObj.setAttribute('id','alertbgDiv');    
		bgObj.style.position="absolute";    
		bgObj.style.top="0";    
		bgObj.style.background="#E8E8E8";    
		bgObj.style.filter="progid:DXImageTransform.Microsoft.Alpha(style=3,opacity=25,finishOpacity=75";    
		bgObj.style.opacity="0.6";    
		bgObj.style.left="0";    
		bgObj.style.width = sWidth + "px";    
		bgObj.style.height = sHeight + "px";    
		bgObj.style.zIndex = "10000";    
		document.body.appendChild(bgObj);    
		//创建提示窗口的div    
		var msgObj = document.createElement("div")    
		msgObj.setAttribute("id","alertmsgDiv");    
		msgObj.setAttribute("align","center");    
		msgObj.style.background="white";    
		msgObj.style.border="1px solid " + bordercolor;    
		msgObj.style.position = "absolute";    
		msgObj.style.left = "50%";    
		msgObj.style.font="12px/1.6em Verdana, Geneva, Arial, Helvetica, sans-serif";    
		//窗口距离左侧和顶端的距离     
		msgObj.style.marginLeft = "-225px";    
		//窗口被卷去的高+（屏幕可用工作区高/2）-150    
		msgObj.style.top = document.body.scrollTop+(window.screen.availHeight/2)-150 +"px";    
		msgObj.style.width = msgw + "px";    
		msgObj.style.height = msgh + "px";    
		msgObj.style.textAlign = "center";    
		msgObj.style.lineHeight ="25px";    
		msgObj.style.zIndex = "10001";    
		document.body.appendChild(msgObj);    
		//提示信息标题    
		var title=document.createElement("h4");    
		title.setAttribute("id","alertmsgTitle");    
		title.setAttribute("align","left");    
		title.style.margin="0";    
		title.style.padding="3px";    
		title.style.background = bordercolor;    
		title.style.filter="progid:DXImageTransform.Microsoft.Alpha(startX=20, startY=20, finishX=100, finishY=100,style=1,opacity=75,finishOpacity=100);";    
		title.style.opacity="0.75";    
		title.style.border="1px solid " + bordercolor;    
		title.style.height="18px";    
		title.style.font="12px Verdana, Geneva, Arial, Helvetica, sans-serif";    
		title.style.color="white";    
		title.innerHTML="提示信息";    
		document.getElementById("alertmsgDiv").appendChild(title);    
		//提示信息    
		var txt = document.createElement("p");    
		txt.setAttribute("id","msgTxt");    
		txt.style.margin="16px 0";    
		txt.innerHTML = str;    
		document.getElementById("alertmsgDiv").appendChild(txt);    
		//设置关闭时间    
		window.setTimeout("closewin()",2000);
	}    
	function closewin() {    
		document.body.removeChild(document.getElementById("alertbgDiv"));    
		document.getElementById("alertmsgDiv").removeChild(document.getElementById("alertmsgTitle"));    
		document.body.removeChild(document.getElementById("alertmsgDiv"));    
	}  
</script>  
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
<form id="form1" name="form1" method="POST" action="<%=MM_editAction%>">
  <table width="640" align="center" bgcolor="#CCCCCC"  cellpadding="5">
    <tr class="beijing">
      <td class="biaoti" align="center">签写留言</td>
    </tr>
    <tr>
      <td align="right"><a href="index.asp"><img src="image/view.jpg"  /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="login.asp"><img src="image/admin.jpg"  /></a></td>
    </tr>
    <tr>
      <td><table width="100%" border="0">
        <tr>
          <td width="18%">姓名：</td>
          <td width="39%"><span id="sprytextfield1">
            <label for="Name"></label>
            <input type="text" name="Name" id="Name" />
            <span class="textfieldRequiredMsg">需要提供一个值。</span></span></td>
          <td width="18%">邮箱：</td>
          <td width="25%"><span id="sprytextfield3">
          <label for="Email"></label>
          <input type="text" name="Email" id="Email" />
          <span class="textfieldRequiredMsg">需要提供一个值。</span><span class="textfieldInvalidFormatMsg">格式无效。</span></span></td>
        </tr>
        <tr>
          <td>主页：</td>
          <td><span id="sprytextfield2">
            <label for="Homepage"></label>
            <input type="text" name="Homepage" id="Homepage" />
            <span class="textfieldRequiredMsg">需要提供一个值。</span></span></td>
          <td>QQ：</td>
          <td><span id="sprytextfield4">
          <label for="QQ"></label>
          <input type="text" name="QQ" id="QQ" />
          <span class="textfieldRequiredMsg">需要提供一个值。</span><span class="textfieldInvalidFormatMsg">格式无效。</span></span></td>
        </tr>
        <tr>
          <td>来自：</td>
          <td><label for="fromwhere"></label>
            <input type="text" name="fromwhere" id="fromwhere" /></td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>头像</td>
          <td colspan="3"><table width="100%" border="0">
            <tr>
              <td><img src="image/01.gif"/>
                <input type="radio" name="ICON" id="ICON" value="01.gif" />
                <label for="ICON"></label></td>
              <td><img src="image/02.gif"/>
                <input type="radio" name="ICON" id="ICON" value="02.gif" /></td>
              <td><img src="image/03.gif"/>
                <input type="radio" name="ICON" id="ICON" value="03.gif" /></td>
              <td><img src="image/04.gif"/>
                <input type="radio" name="ICON" id="ICON" value="04.gif" /></td>
              <td><img src="image/05.gif"/>
                <input type="radio" name="ICON" id="ICON" value="05.gif" /></td>
            </tr>
            <tr>
              <td><img src="image/06.gif"/>
                <input type="radio" name="ICON" id="ICON" value="06.gif" /></td>
              <td><img src="image/07.gif"/>
                <input type="radio" name="ICON" id="ICON" value="07.gif" /></td>
              <td><img src="image/08.gif"/>
                <input type="radio" name="ICON" id="ICON" value="08.gif" /></td>
              <td><img src="image/09.gif"/>
                <input type="radio" name="ICON" id="ICON" value="09.gif" /></td>
              <td><img src="image/10.gif"/>
                <input type="radio" name="ICON" id="ICON" value="10.gif" /></td>
            </tr>
            <tr>
              <td><img src="image/11.gif"/>
                <input type="radio" name="ICON" id="ICON" value="11.gif" /></td>
              <td><img src="image/12.gif"/>
                <input type="radio" name="ICON" id="ICON" value="12.gif" /></td>
              <td><img src="image/13.gif"/>
                <input type="radio" name="ICON" id="ICON" value="13.gif" /></td>
              <td><img src="image/14.gif"/>
                <input type="radio" name="ICON" id="ICON" value="14.gif" /></td>
              <td><img src="image/15.gif"/>
                <input type="radio" name="ICON" id="ICON" value="15.gif" /></td>
            </tr>
            <tr>
              <td><img src="image/16.gif"/>
                <input type="radio" name="ICON" id="ICON" value="16.gif" /></td>
              <td><img src="image/17.gif"/>
                <input type="radio" name="ICON" id="ICON" value="17.gif" /></td>
              <td><img src="image/18.gif"/>
                <input type="radio" name="ICON" id="ICON" value="18.gif" /></td>
              <td><img src="image/19.gif"/>
                <input type="radio" name="ICON" id="ICON" value="19.gif" /></td>
              <td><img src="image/20.gif"/>
                <input type="radio" name="ICON" id="ICON" value="20.gif" /></td>
            </tr>
          </table></td>
          </tr>
        <tr>
          <td>留言：</td>
          <td colspan="3"><span id="sprytextarea1">
            <label for="Content"></label>
            <textarea name="Content" id="Content" cols="60" rows="5"></textarea>
            <span class="textareaRequiredMsg">需要提供一个值。</span></span></td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td align="center"><input type ="submit" name ="submit" value ="提交留言"/>&nbsp; &nbsp;&nbsp;&nbsp;
 <input type ="reset" name ="reset" value ="重写留言"  onClick="Alert('信息已全部清空，请重新填写')"/>
 <input name="Date" type="hidden" id="Date" value="<%=Date%>" /></td>
    </tr>
    <tr>
      <td class="zhengwen"  align="center">版权所有17教育技术学</td>
    </tr>
  </table>
  <input type="hidden" name="MM_insert" value="form1" />
</form>
<script type="text/javascript">
var sprytextarea1 = new Spry.Widget.ValidationTextarea("sprytextarea1");
var sprytextfield1 = new Spry.Widget.ValidationTextField("sprytextfield1");
var sprytextfield2 = new Spry.Widget.ValidationTextField("sprytextfield2");
var sprytextfield3 = new Spry.Widget.ValidationTextField("sprytextfield3", "email");
var sprytextfield4 = new Spry.Widget.ValidationTextField("sprytextfield4", "integer");
</script>
</body>
</html>
