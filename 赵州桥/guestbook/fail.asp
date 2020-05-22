<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>登录失败</title>
<link href="style.css" rel="stylesheet" type="text/css" />
<style type="text/css">
img {
	height: 50px;
	width: 50px;
}
</style>
<link href="../SpryAssets/SpryValidationTextarea.css" rel="stylesheet" type="text/css" />
<link href="../SpryAssets/SpryValidationTextField.css" rel="stylesheet" type="text/css" />
<meta http-equiv="refresh" content="3;URL=login.asp" />
</head>


<body>

  <table width="640"align="center" bgcolor="#CCCCCC"  cellpadding="5">
    <tr class="beijing">
      <td class="biaoti" align="center">登录失败</td>
    </tr>
    <tr>
      <td align="right"><a href="index.asp"><img src="image/view.jpg"  /></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="insert.asp"><img src="image/add.jpg"  /></a></td>
    </tr>
    <tr align="center">
      <td class="hongzi">账号或密码错误错误，登录失败，请重新登录！</td>
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
