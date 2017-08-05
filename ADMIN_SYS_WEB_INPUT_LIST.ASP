<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->

<%
'|---------------------------------------------------------------------------------------
'|---------------------------   必要参数,接收表单参数、数据库参数    --------------------|
'                                                                                        |
dim db_tabl,db_key                            '数据库参数
dim del_id,del_back,page_size,list_num,list_conn_num        '删除返回的数据


'以上的是配置信息 ————————
    db_tabl = "sys_web_input"                          '当前操作的表
	db_key="id"                           '主键

	page_size=10                              '每页显示数目
	
	del_id=request.querystring("del_id")      '删除的id值
	
sub show_msg(str)
    response.write("<script>alert('"&str&"');</script>")
end sub
	
	
	
'--->>判断并删除
if del_id<>"" and isnumeric(del_id) then
set list_del_conn=server.createobject("adodb.recordset")

    list_del_check_str="select * from " & db_tabl & "  where "&db_key&"=" & int(del_id)
	list_del_conn_str="delete * from " & db_tabl & "  where "&db_key&"=" & int(del_id)
	list_del_conn.open list_del_check_str,connstr,1,3
	
	if not list_del_conn.eof then
	   list_del_conn.close
	   list_del_conn.open list_del_conn_str,connstr,1,3
	   list_del_conn.update
       del_back="true"
	else
	   list_del_conn.close
	end if
set	list_del_conn=nothing
end if
	
	

'------->读取数据
set list_conn=server.createobject("adodb.recordset")
    list_conn_str="select * from " & db_tabl & " order by "&db_key&" asc"
	list_conn.open list_conn_str,connstr,1,1
%>
<link href="fanren_style/style.css" rel="stylesheet" type="text/css">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
  <tr>
    <td height="25" colspan="3" bgcolor="#ffffff">&nbsp; <font color="#135294"><strong>输入框名称</strong></font></td>
  </tr>
  <tr>
  
  
<td class="forumrow">&nbsp;&nbsp;输入框名称</td><td height="25" colspan="3" align="center" class="forumrow">管理</td>
  </tr>
<%do while not  list_conn.eof%>
<tr>
<td class="forumrow">&nbsp;·<a href="admin_<%=db_tabl%>_edit.asp?page=<%=page%>&edit_type=edit&edit_id=<%=list_conn(db_key)%>"><%=list_conn("web_input_name")%></a></td><td width="60" align="center" class="forumrow">
<a href="?page=<%=page%>&del_id=<%=list_conn(db_key)%>" onclick="return confirm('是否确定删除?');">删除</a>
</td>
<td width="60" align="center" class="forumrow">
<a href="admin_<%=db_tabl%>_edit.asp?page=<%=page%>&edit_type=edit&edit_id=<%=list_conn(db_key)%>">修改</a></td>
</tr>
<%
    list_conn.movenext
    loop
    list_conn.close
set list_conn =nothing
%>  

<tr>
  <td colspan="3">
 </td>
  </tr>
</table>


<%
if  del_back="true" then
call show_msg("已经成功删除...")
end if
%>

