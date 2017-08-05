<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->
<%
'|---------------------------------------------------------------------------------------
'|---------------------------   必要参数,接收表单参数、数据库参数    --------------------|
'                                                                                        |
dim db_tabl,db_key                            '数据库参数
dim del_id,del_back,page_size,list_num,list_conn_num        '删除返回的数据


'以上的是配置信息 ————————
    db_tabl = "sys_web_code_type"                          '当前操作的表
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
    list_conn_str="select * from " & db_tabl & " order by "&db_key&" desc"
	list_conn.open list_conn_str,connstr,1,3
'------->数据分页

if not list_conn.eof then
   list_conn.pagesize=page_size
if request.querystring("page")<>"" then
   page=cint(request.querystring("page"))   
   if page<1 then page=1
   if page>list_conn.pagecount then page=list_conn.pagecount
else
   page=1
end if
   list_conn.absolutepage=page
end if

'------->分页必须的变量
page=page                      '当前页数
list_num=3                     '
'frist_page=1                   ' 首页
last_page=list_conn.pagecount  ' 最后页
%>
<link href="fanren_style/style.css" rel="stylesheet" type="text/css">
<script type="text/javascript">
<!--
function mm_openbrwindow(theurl,winname,features) { //v2.0
  window.open(theurl,winname,features);
}
//-->
</script>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
  <tr>
    <td height="25" colspan="4" bgcolor="#ffffff">&nbsp;&nbsp;<strong><font color="#135294">代码分类</font></strong></td>
  </tr>
  
  <tr>

<td width="616" class="forumrow">&nbsp;&nbsp;分类名称</td>

    
    <td height="25" colspan="3" align="center" class="forumrow">管理&nbsp;|| <a href="#" onclick="mm_openbrwindow('admin_sys_web_code_type_edit.asp?edit_type=add','codeadd','width=350,height=120')">添加</a></td>
  </tr>
<%
for list_conn_num=1 to page_size
if  list_conn.eof or list_conn.bof then exit for
%>
<tr>


<td class="forumrow">&nbsp;&nbsp;·<a href="admin_sys_web_code_list.asp?code_type_id=<%=list_conn("id")%>"><%=list_conn("code_type_name")%></a></td>


<td width="35" align="center" class="forumrow">
<a href="?page=<%=page%>&del_id=<%=list_conn(db_key)%>" onclick="return confirm('是否确定删除?');">删除</a></td>
<td width="35" align="center" class="forumrow">
<a href="#" onclick="mm_openbrwindow('admin_sys_web_code_type_edit.asp?page=<%=page%>&edit_type=edit&edit_id=<%=list_conn(db_key)%>','codeadd','width=350,height=120')">
修改</a></td>
</tr>

<%
    list_conn.movenext
     next
    list_conn.close
set list_conn =nothing
%>  

<tr>
  <td colspan="4" class="forumrow"><!--#include file="admin_all_pagination.asp"--></td>
</tr>
</table>


<%
if  del_back="true" then
call show_msg("已经成功删除...")
end if
%>

