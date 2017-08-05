<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->
<!--#include file="fanren_include/f_funtion.asp"-->


<%
'|---------------------------------------------------------------------------------------
'|---------------------------   必要参数,接收表单参数、数据库参数    --------------------|
'                                                                                        |
dim db_tabl,db_key                            '数据库参数
dim del_id,del_back,page_size,list_num,list_conn_num        '删除返回的数据
code_type_id =request.querystring("code_type_id")

if code_type_id="" or isnumeric(code_type_id) =false then
   response.write("参数有误！")
   response.end()
end if


'以上的是配置信息 ————————
    db_tabl = "sys_web_code"                          '当前操作的表
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
    list_conn_str="select * from " & db_tabl & " where code_type_id="&int(code_type_id)&" order by "&db_key&" desc"
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
    <td height="25" colspan="6" bgcolor="#ffffff">&nbsp;&nbsp; <font color="#135294"><strong>分类管理</strong></font></td>
  </tr>
  
  <tr>
<td width="2" class="forumrow">&nbsp;</td>
<td width="261" class="forumrow">代码名称</td>
<td width="335" class="forumrow">添加时间</td>

    
    <td height="25" colspan="3" align="center" class="forumrow">管理&nbsp;|| <a href="javascript:mm_openbrwindow('admin_sys_web_code_edit.asp?edit_type=add&code_type_id=<%=code_type_id%>','codeadd','width=550,height=275')">添加</a></td>
  </tr>
<%
for list_conn_num=1 to page_size
if  list_conn.eof or list_conn.bof then exit for
%>
<tr>
  <td width="2" class="forumrow">&nbsp;</td>

<td class="forumrow">
<%
my_code_type_id=list_conn("code_type_id")
call get_code_type(1,my_code_type_id)%>
&nbsp;<%=list_conn("code_name")%></td>
<td class="forumrow">
<%=list_conn("code_add_time")%></td>


<td width="38" align="center" class="forumrow">
<a href="?page=<%=page%>&del_id=<%=list_conn(db_key)%>" onclick="return confirm('是否确定删除?');">删除</a></td>
<td width="38" align="center" class="forumrow">
<a href="javascript:mm_openbrwindow('admin_<%=db_tabl%>_edit.asp?page=<%=page%>&edit_type=edit&edit_id=<%=list_conn(db_key)%>&code_type_id=<%=code_type_id%>','codeadd','width=550,height=275')">
修改</a></td>
</tr>

<%
    list_conn.movenext
     next
    list_conn.close
set list_conn =nothing
%>  
<tr>
  <td colspan="6" class="forumrow">
<!--#include file="admin_all_pagination.asp"--></td>
  </tr>
</table>

<%
if  del_back="true" then
call show_msg("已经成功删除...")
end if
%>

