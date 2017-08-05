<%@language="vbscript" codepage="65001"%>
<% response.codepage=65001%>
<% response.charset="utf-8" %>
<!--#include file="fanren_include/f_md5.asp"-->
<!--#include file="admin_conn.asp"-->


<%'筛选，判断是登录还是退出
g_event=request.querystring("event")
%>

<%
select case g_event
case ""
'=======================================

'检测是否从注册页面进入,若存在问题则返回
if session("checkcode")="" then
   response.redirect("admin_login.asp?id=8")
   response.end()
end if
%>

<%
'定义过滤非法字符函数
function html_str(str)
         r_str=str
		 c_str="',<,>,|,#,%,?,=,;"
		 str_arr=split(c_str,",")
		 for each e_str in str_arr
		     str=replace(str,e_str,"")
		 next
		 html_str=str
		 
		 if str<>r_str then
		    response.redirect("admin_login.asp?id=2")
			response.end()
		 end if
end function
%>


<%
'接收表单信息
login_name=request.form("user")
login_pass=request.form("pass")
login_num=request.form("char")

'检测非法字符
call html_str(login_name)
call html_str(login_pass)
call html_str(login_num)
%>


<%
'检测是否从注册页面进入,若存在问题则返回
if login_num="" then
   response.write("<script>alert('请输入验证码！');history.go(-1);</script>")
   response.end()
end if
%>


<%
'验证码
if int(login_num)<>int(session("checkcode")) then
   response.redirect("admin_login.asp?id=3")
   response.end()
end if
%>


<%
'验帐号
if len(login_name)=0 then
   response.write("<script>alert('请输入[登录帐号]');history.go(-1);</script>")
else

set login_user=server.createobject("adodb.recordset")
    login_user_str="select * from sys_fanren_admin where admin_name='" & cstr(login_name) & "'"
    login_user.open login_user_str,conns,1,3

if not login_user.eof then
       g_login_pass=login_user("admin_pass")
       g_login_id=login_user("id")

   
   if md5(login_pass)=cstr(g_login_pass) then
      session("admin_id")=g_login_id
      session("admin_name")=login_name
	  
      response.write("<script>alert(' \n => [互意网络] 提示! \n\n => [时间] "&now()&" \n\n => [操作] 登录 \n\n => [用户] "&login_name&" \n\n => [返回] 成功登录！');window.location.href='admin_default.asp';</script>")
response.end()
	  
	  else
	  response.redirect("admin_login.asp?id=1")
      response.end()
   end if
   
else
      response.redirect("admin_login.asp?id=1")
      response.end()
end if

login_user.close
set login_user=nothing
end if
%>


<%
'=======================================
case "out"

      session.abandon()   '释放所以的session值
      response.redirect("admin_login.asp?id=9")
      response.end()

end select
%>
