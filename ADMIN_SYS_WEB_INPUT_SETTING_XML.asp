<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->
<!--#include file="fanren_include/f_funtion.asp"-->
<?xml version="1.0" encoding="utf-8"?>
<files id="1" name="1">
<%
'数据库连接===================================

dim sys_conns,sys_connstr,sys_mdb

    on error resume next
    sys_connstr="driver=microsoft access driver (*.mdb);dbq=" + session("web_db")


   sys_table=request.querystring("tables_id")
if sys_table<>"" then
	
    '------------------------ 	
set sys_conns=server.createobject("adodb.connection") 
    sys_conns.open sys_connstr


'\\\\\\\\\\\\\判断数据库文件是否存在\\\\\\\\\\\\\\\\\\\\\\
if err then
   err.clear
set sys_connstr = nothing
   response.write "<script>alert('数据库文件已不存在，请重选...');parent.location.href='admin_default.asp';</script>"
   
   session("web_no.")          = ""
   response.cookies("web_no.") = ""
   
   response.end
end if
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'------------------------  
%><%
set show_filed_conn = server.createobject("adodb.recordset")
    show_filed_conn_str = "select * from " & sys_table
    show_filed_conn.open show_filed_conn_str,sys_connstr,0,1
	
	for i=0 to show_filed_conn.fields.count-1
%>
<x><%=show_filed_conn.fields.item(i).name%></x>
<%
	next
	
	show_filed_conn.close
set	show_filed_conn=nothing

end if
%> 
</files>