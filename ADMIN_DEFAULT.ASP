<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->
<!--#include file="fanren_include/f_funtion.asp"-->
<html>
<head>
<title>欢迎使用-<%=sitename%>-后台管理</title>
<link href="fanren_style/style2.css" rel="stylesheet" type="text/css" >
<style type="text/css">
<!--
.style1 {color: #cccccc}
-->
</style>
<meta http-equiv="content-type" content="text/html; charset=utf-8"></head>

<script>
var status = 1;
//var menus = new dvmenucls;
//document.onclick=menus.clear;
function switchsysbar(){
     if (1 == window.status){
		  window.status = 0;
          switchpoint.innerhtml = '<img src="fanren_images/left.gif">';
          document.all("frmtitle").style.display="none"
     }
     else{
		  window.status = 1;
          switchpoint.innerhtml = '<img src="fanren_images/right.gif">';
          document.all("frmtitle").style.display=""
     }
}
</script>

<script>
function out(){
if(confirm("是否退出网站管理系统？")){
window.location.href="admin_login_check.asp?event=out";
return true;
}
}


</script>




<body>
<!--导航部分开始-->
<div class="top_table">
<div class="top_table_leftbg">
	<div class="system_logo"></div>
  <div class="system_lano">
	<ul id="nav">
<li><a target="_top" href="javascript:out();" class="nav_product">退出系统</a></li>
<li><a target="main" href="admin_edit.asp" class="nav_product">管理账号</a></li>
<li><a target="main" href="admin_info.asp" class="nav_product">网站信息</a></li>
<li><a target="main" href="admin_sys_web_template_list.asp?default=0" class="nav_product">默认模板</a></li>
<%if session("web_no.")<>"" then%>
<li><a target="left" href="admin_left.asp" class="nav_product">管理菜单</a></li>
<li><a target="left" href="admin_left.asp?show_db=0" class="nav_product">选数据表</a></li>
<li><a href="?reget_web_db=yes" class="nav_product" onclick="return confirm('是否确定重选数据库?');">重选数据库</a></li>
<%else%>
<li><a target="main" href="admin_info.asp" class="nav_product">选数据库</a></li>
<%end if%>
<li>
  <table border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td valign="bottom"><span class="style1">欢迎您： <%=session("admin_name")%></span></td>
      <td valign="bottom"> <img src="fanren_images/arrow.gif" alt="" align="absmiddle" >&nbsp;</td>
      <td valign="bottom" style="text-align:right; color: #135294; "><a href="http://www.sznt.net" target='_blank'> 客服网站</a> | <a href="admin_default.asp" target='_top'> 后台首页</a></td>
      <td width="8" align="right" valign="top" >&nbsp;</td>
      </tr>
  </table>
</li>
	  </ul>
	</div>
</div>
</div>
<!--导航部分结束-->
<table width="100%" height="93%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td id=frmtitle valign=top name="fmtitle" class="main_left" bgcolor="#000000"><iframe frameborder=0 id=frmleft name="left" src="admin_left.asp" class="left_iframe" allowtransparency="true"></iframe></td>
    <td width="1%" valign="top" bgcolor="#000000"><table border="0" cellpadding="0" cellspacing="0" height="100%">
      <tbody>
        <tr>
          <td width="13" align="center" style="height: 100%" onclick="switchsysbar()"><span class="navpoint" id="switchpoint" title="关闭/打开左栏"><img src="fanren_images/right.gif" width="9" height="41" ></span> </td>
        </tr>
      </tbody>
    </table></td>
    <td width="100%" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="15" bgcolor="#000000">&nbsp;</td>
      </tr>
    </table>
        <iframe frameborder=0 id=frmright name="main" scrolling=yes src="admin_info.asp" class="main_iframe"></iframe>

	</td> 
  </tr>
</table>
</body>