<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->
<!--#include file="fanren_include/f_md5.asp"-->
<%
'修改管理员的代码
   admin_edit_form=request.form("admin_edit_form")
if admin_edit_form="edit" and session("admin_id")<>"" and isnumeric(session("admin_id")) then
admin_name=request.form("admin_name")
admin_pass=request.form("admin_pass")

set admin_edit_conn=server.createobject("adodb.recordset")
	admin_edit_conn_str="select * from sys_fanren_admin where id="&int(session("admin_id"))
    admin_edit_conn.open admin_edit_conn_str,conns,1,3
	
	   admin_edit_back="false"
	if not admin_edit_conn.eof and not admin_edit_conn.bof then
	   if admin_edit_conn("admin_pass")<>"" then
	      admin_edit_conn("admin_pass")=md5(admin_pass)
		  'admin_edit_conn("admin_name")=admin_name
	      'session("admin_name")=admin_name
	      admin_edit_conn.update
	      admin_edit_back="true"
		  else
		  admin_edit_back="false"
	   end if
	end if
	
	admin_edit_conn.close
set	admin_edit_conn=nothing
end if
%>


<link href="fanren_style/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body,table,tr,td{font-size:12px;}
.style1 {font-size: 12px}
.style2 {
	color: #ff0000;
	font-size: 12px;
}
.style3 {
	color: #333333;
	font-size: 14px;
}
.style5 {color: #ff0000; font-size: 14px; }
-->
</style>


<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
  <tr>
    <td height="25" bgcolor="#ffffff">&nbsp;&nbsp; <font color="#135294"><strong>管理帐号管理</strong></font></td>
  </tr>
  <tr>
    <td height="100" align="center" valign="middle" class="forumrow">
<form action="" name="edit_pwd_form" method="post" style="margin:0px;" onsubmit="return check_form()">
     <table width="500" border="0" align="center" cellpadding="0" cellspacing="1">
      <tr>
        <td width="200" height="25" align="right"><span class="style1">管理帐号：</span></td>
        <td width="23%"><label>
<input disabled="disabled" name="admin_name" type="text" id="admin_name" value='<%=session("admin_name")%>'/>
        </label></td>
        <td width="41%">&nbsp;</td>
      </tr>
      <tr>
        <td height="25" align="right"><span class="style1">修改密码：</span></td>
        <td><input type="password" name="admin_pass" id="admin_pass" /></td>
        <td><span class="style2"> &nbsp;* 长度为 5 ~ 12 位</span></td>
      </tr>
      <tr>
        <td height="25" align="right"><span class="style1">确认密码：</span></td>
        <td><input type="password" name="admin_pass1" id="admin_pass1" /></td>
        <td><span class="style2">&nbsp;* 长度为 5 ~ 12 位</span></td>
      </tr>
      <tr>
        <td height="25">&nbsp;</td>
        <td height="25"><label>
          <input type="reset" name="button2" id="button2" value="重置" />
          <input type="submit" name="button" id="button" value="确认修改" />
        </label></td>
        <td height="25">&nbsp;</td>
      </tr>
    </table>
    <input type="hidden" name="admin_edit_form" value="edit" />
    </form>    

<script language="javascript">
function check_form()
{

if(document.edit_pwd_form.admin_pass.value.length<5)
{
alert("不能少于5个字符！");
document.edit_pwd_form.admin_pass.focus();
return false;
}

if(document.edit_pwd_form.admin_pass.value.length>12)
{
alert("不能超过12个字符！");
document.edit_pwd_form.admin_pass.focus();
return false;
}

if(document.edit_pwd_form.admin_pass.value==document.edit_pwd_form.admin_pass1.value)
{
} else {

alert("两次输入的密码不一致...");
document.edit_pwd_form.admin_pass.focus();
return false;

}

}
</script>  

<%
if admin_edit_back="true" then
response.write("<script>alert('已成功修改密码！');</script>")
elseif admin_edit_back="false" then
response.write("<script>alert('密码修改失败！');</script>")
end if
%>

</td>
  </tr>
</table>
