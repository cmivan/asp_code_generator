<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->
<!--#include file="fanren_include/f_funtion.asp"-->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8">
<title>后台管理</title>
<link href="fanren_style/admin_left.css" rel=stylesheet type=text/css>
<link href="fanren_style/style2.css" rel=stylesheet type=text/css>
<link href="fanren_style/style_left.css" rel=stylesheet type=text/css>
<style type="text/css">
<!--
.style1 {
	font-size: 14px;
	font-weight: bold;
	color: #999999;
}

-->
</style></head>
<script lanuage="javascript">
var tt='start';
var ii='start';
function turnit(ss,bb) {

  if (ss.style.display=="none") {
    if(tt!='start') tt.style.display="none";
    if(ii!='start') ii.src="";
    ss.style.display="";
    tt=ss;
    ii=bb;
    bb.src="";
  }
  else {
    ss.style.display="none"; 
    bb.src="";
  }
}
</script>
</head>
<body>
<table width=181 border='0' cellpadding=0 cellspacing=0 bgcolor="#25a2dd">
  <tr>
    <td width="25" height=1 valign=top bgcolor="#000000">&nbsp;</td>
  </tr>
</table>
<table width="181" height="100%" border="0" cellpadding="0" cellspacing="0" background="fanren_images/left_bg.gif">
  <tr>
    <td height="40" background="fanren_images/left_menu_bg.gif"><span class="style1">&nbsp; 管理菜单</span></td>
  </tr>
  <tr>
    <td valign="top">
	<table width=181 border="0" cellpadding=0 cellspacing=0>
      <tr>
        <td >
		<table width="181" border="0" cellpadding="0" cellspacing="0" class="alpha">
            <tr>
              <td valign="top" class="menu" id="menubar">


<%
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\读取数据库内容，表、字段\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
if request.querystring("show_db")<>"" and session("web_no.")<>"" and session("web_db")<>"" and session("web_db_fullpath")<>"" then
%>

<div id="my_db_link">
<%call get_web_db_table("list") '///读取数据库表段%>
</div>




<%else%>

		    <table cellspacing="0" cellpadding="0" width=100% >
		      <tr>
			   <td  id=tag1 style='cursor: hand' onclick=turnit(content1,tag1);  language=jscript ><b class="main_left_title"><img src="fanren_images/bullet.gif" border=0>&nbsp;设置面板</b> </td>
		      </tr>
		      <tr>
			    <td id=content1 style="display: none" align="center">	

			
<li><a target="main" href="admin_sys_web_path_list.asp">路径管理</a></li>
<li><a target="main" href="admin_sys_web_template_list.asp">模板管理</a></li>
<li><a target="main" href="admin_sys_web_input_list.asp">输入框设置</a></li>
<li><a target="main" href="admin_sys_web_input_limit_list.asp">输入限制设置</a></li>
<br>
</td>
		     </tr>
	        </table>	<br>
	        <table cellspacing="0" cellpadding="0" width=100%  >
		      <tr>
			   <td  id=tag2 style='cursor: hand' onclick=turnit(content2,tag2);  language=jscript ><b class="main_left_title"><img src=fanren_images/bullet.gif border=0>&nbsp;模板管理</b> </td>
		      </tr>
		      <tr>
			    <td id=content2 style="display: none" align="center" >
<%call get_tmpl_name("sys_web_template","web_name",0,"li")%>
<br>
			   </td>
		     </tr>
	        </table>	<br> 
			<table cellspacing="0" cellpadding="0" width=100% >
		      <tr>
			   <td  id=tag3 style='cursor: hand' onclick=turnit(content3,tag3);  language=jscript ><b class="main_left_title"><img src=fanren_images/bullet.gif border=0>&nbsp;代码收集</b> </td>
		      </tr>
		      <tr>
			    <td id=content3 style="display: none" align="center">				

<%
'------->读取数据
set list_conn=server.createobject("adodb.recordset")
    list_conn_str="select * from sys_web_code_type order by id desc"
	list_conn.open list_conn_str,connstr,1,1
'------->数据分页
    do while not list_conn.eof
%>
<li><a target="main" href="admin_sys_web_code_list.asp?code_type_id=<%=list_conn("id")%>"><%=list_conn("code_type_name")%></a>
</li>
<%
    list_conn.movenext
    loop
    list_conn.close
set list_conn =nothing
%>

<li><a target="main" href="admin_sys_web_code_type_list.asp" >分类管理</a>
</li>
			   <br>
			   </td>
		     </tr>
	        </table>	<br>
	        <table cellspacing="0" cellpadding="0" width=100% >
		      <tr>
			   <td  id=tag4 style='cursor: hand' onclick=turnit(content4,tag4);  language=jscript ><b class="main_left_title"><img src=fanren_images/bullet.gif border=0>&nbsp;用户管理</b> </td>
		      </tr>
		      <tr>
			    <td id=content4 style="display: none" align="center">				
			    <li><a target="main" href="#">用户信息</a></li>
           	   <br>
			   </td>
		     </tr>
	        </table>	<br>
	        <table cellspacing="0" cellpadding="0" width=100% >
		      <tr>
			   <td  id=tag5 style='cursor: hand' onclick=turnit(content5,tag5);  language=jscript ><b class="main_left_title"><img src=fanren_images/bullet.gif border=0>&nbsp;代码收集</b> </td>
		      </tr>
		      <tr>
			    <td id=content5 style="display: none" align="center">				
			    <li><a target="main" href="#" >asp代码</a></li>
		        <li><a target="main" href="#" >php代码</a></li>
			   </td>
		     </tr>
	        </table>	<br> 
			<table cellspacing="0" cellpadding="0" width=100% >
		      <tr>
			   <td  id=tag6 style='cursor: hand' onclick=turnit(content6,tag6);  language=jscript ><b class="main_left_title"><img src=fanren_images/bullet.gif border=0>&nbsp;数据管理</b> </td>
		      </tr>
		      <tr>
			    <td id=content6 style="display: none" align="center">				
                <li><a target="main" href="admin_stat.asp">数据统计</a></li>
                <br>
				</td>
		     </tr>
	        </table>	<br> 
			<table cellspacing="0" cellpadding="0" width=100% >
<tr>
<td  id=tag7 style='cursor: hand' onclick=turnit(content7,tag7);  language=jscript ><b class="main_left_title">
<img src=fanren_images/bullet.gif border=0>&nbsp;系统相关</b> </td>
		      </tr>
		      <tr>
<td id=content7 style="display: none" align="center">				
<li><a target="main" href="aspcheck.asp" >服务器情况</a></li>
         	  <br>
			   </td>
		     </tr>
	        </table>	
<%
end if
%>      
	 </td>
		</tr>
        </table>
		</td>
      </tr>
    </table>
	</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>