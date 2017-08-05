<%@language="vbscript" codepage="65001"%>
<% response.codepage=65001%>
<% response.charset="utf-8" %>
<!doctype html public "-//w3c//dtd xhtml 1.0 transitional//en" "http://www.w3.org/tr/xhtml1/dtd/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file ="fanren_include/f_info.asp" -->
<meta name="author" content="fanren,619835864@qq.com " />
<meta name="copyrigth" content="深圳新格构科技有限公司版权所有-www.sznt.net" />
<meta name="description" content="深圳专业设计研究机构，专注视觉,人机交互体验-www.sznt.net" />
<meta name="keywords" content="深圳设计,品牌策划,vi设计,网站设计,软件开发," />
<title>trellis 网站后台管理</title>
<link href="fanren_style/admin_login.css" rel="stylesheet" type="text/css" />
</head>
<script language="javascript" type="text/javascript">
<!--
function frontpage_form1_validator(theform)
{

  if (theform.user.value == "")
  {
    alert("请在 用户名 域中输入值。");
    theform.user.focus();
    return (false);
  }

  if (theform.user.value.length < 2)
  {
    alert("在 用户名 域中，请至少输入 2 个字符。");
    theform.user.focus();
    return (false);
  }

  if (theform.user.value.length > 20)
  {
    alert("在 用户名 域中，请最多输入 20 个字符。");
    theform.user.focus();
    return (false);
  }

  if (theform.pass.value == "")
  {
    alert("请在 密码 域中输入值。");
    theform.pass.focus();
    return (false);
  }

  if (theform.pass.value.length < 5)
  {
    alert("在 密码 域中，请至少输入 5 个字符。");
    theform.pass.focus();
    return (false);
  }

  if (theform.pass.value.length > 20)
  {
    alert("在 密码 域中，请最多输入 20 个字符。");
    theform.pass.focus();
    return (false);
  }

  if (theform.artyzm.value == "")
  {
    alert("请在 验证码 域中输入值。");
    theform.artyzm.focus();
    return (false);
  }

  if (theform.artyzm.value.length < 4)
  {
    alert("在 验证码 域中，请至少输入 4 个字符。");
    theform.artyzm.focus();
    return (false);
  }

  if (theform.artyzm.value.length > 6)
  {
    alert("在 验证码 域中，请最多输入 4 个字符。");
    theform.artyzm.focus();
    return (false);
  }

  var checkok = "0123456789abcdefghijklmnopqrstuvwxyz-";
  var checkstr = theform.artyzm.value;
  var allvalid = true;
  var validgroups = true;
  var decpoints = 0;
  var allnum = "";
  for (i = 0;  i < checkstr.length;  i++)
  {
    ch = checkstr.charat(i);
    for (j = 0;  j < checkok.length;  j++)
      if (ch == checkok.charat(j))
        break;
    if (j == checkok.length)
    {
      allvalid = false;
      break;
    }
    allnum += ch;
  }
  if (!allvalid)
  {
    alert("在 验证码 域中，只能输入 数字和字母 字符。");
    theform.artyzm.focus();
    return (false);
  }
  return (true);
}
function refreshimg()
 {
  document.all.checkcode.src='fanren_include/code.gif'; 
  settimeout("refreshimg2()",1100);
   }
   
function refreshimg2()
 {
  document.all.checkcode.src='fanren_code/f_code.asp'; 
   }
//-->
</script>

<body>
<div class="div_center">
<form  form method="post" action="admin_login_check.asp" target=_top onsubmit="return frontpage_form1_validator(this)" language="javascript" style="margin:0;" name="frontpage_form1">
<table width="273" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="34" colspan="2" background="fanren_images/fanren_login/login_title.png">&nbsp;</td>
    </tr>
  <tr>
    <td width="65" height="25"><img src="fanren_images/fanren_login/user.png" width="65" height="20" /></td>
    <td width="208" valign="bottom" class="input_style">
    <input name="user"  type="text" onfocus="this.select();" value="admin"  maxlength="20" /></td>
    </tr>
  <tr>
    <td height="25"><img src="fanren_images/fanren_login/psw.png" width="65" height="19" /></td>
    <td valign="bottom" class="input_style">
    <input name="pass" type="password" value="admin" maxlength="20" /></td>
    </tr>
  <tr>
    <td height="25"><img src="fanren_images/fanren_login/code.png" width="65" height="19" /></td>
    <td valign="bottom" class="input_style"><table width="200" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="159"><input name="char" type="text" id="char" value="<%=session("checkcode")%>" maxlength="6" /></td>
<td width="41">
<img src="fanren_code/f_code.asp" alt="图片看不清？点击换个图片" name="checkcode" width="40" height="10" border="0" id="checkcode" style="cursor:hand;" onclick="javascript:refreshimg();" /></td>
        </tr>
      
    </table></td>
    </tr>
  <tr>
    <td height="28">&nbsp;</td>
    <td valign="bottom"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center"><input type="image" src="fanren_images/fanren_login/login.png" width="68" height="24" /></td>
        <td align="center"><img src="fanren_images/fanren_login/clear.png" width="68" height="24" /></td>
        <td align="center"><img src="fanren_images/fanren_login/forget.png" width="68" height="24" /></td>
      </tr>
      
    </table></td>
    </tr>
</table>
</form>

<table border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="367" height="67" align="center" style="color:#cccccc; font-size:12px;" background="fanren_images/fanren_login/login_bg.png">
    &nbsp;
    <%
			if request("id")="0" then response.write "请填写完整！"
			if request("id")="1" then response.write "用户名或密码错误！"
			if request("id")="2" then response.write "包含非法字符,请核对！"
			if request("id")="3" then response.write "验证码错误"
			if request("id")="8" then response.write "网络超时，或未登录。请重新登录！"
			if request("id")="9" then
			   session.abandon()
			   response.write "正常退出后台管理系统"
			end if
			   
	%></td>
    </tr>
  <tr>
    <td><img src="fanren_images/fanren_login/copy.png" width="367" height="98" border="0" usemap="#map" />
      <map name="map" id="map">
        <area shape="rect" coords="36,10,158,86" href="http://www.sznt.net" target="_blank" />
      </map></td>
  </tr>
</table>
</div>
</body>
</html>
