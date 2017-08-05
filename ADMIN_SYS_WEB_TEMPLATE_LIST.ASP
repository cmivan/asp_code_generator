<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->

<%
'|---------------------------------------------------------------------------------------
'|---------------------------   必要参数,接收表单参数、数据库参数    --------------------|
'                                                                                        |
dim db_tabl,db_key                            '数据库参数
dim del_id,del_back,page_size,list_num,list_conn_num,type_id,default       '删除返回的数据


'以上的是配置信息 ————————
    db_tabl = "sys_web_template"              '当前操作的表
	db_key="id"                               '主键

	page_size=10                              '每页显示数目
	
	del_id=request.querystring("del_id")      '删除的id值
	
	type_id=request.querystring("type_id")
	default=request.querystring("default")
	
sub show_msg(str)
    response.write("<script>alert('"&str&"');</script>")
end sub
	
	
	
'--->>判断并删除
if del_id<>"" and isnumeric(del_id) then
set list_del_conn=server.createobject("adodb.recordset")

    list_del_check_str="select * from " & db_tabl & "  where "&db_key&"=" & int(del_id)
	list_del_conn_str="delete * from " & db_tabl & "  where "&db_key&"=" & int(del_id)
	list_del_conn.open list_del_check_str,connstr,1,1
	
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
    
	
	
	
	if type_id<>"" and isnumeric(type_id)=true then
        list_conn_str="select * from " & db_tabl & " where type_id="&type_id&" order by "&db_key&" asc"
	elseif default<>"" and isnumeric(default)=true then
        list_conn_str="select * from " & db_tabl & " where default=1 order by "&db_key&" asc"
	else
	    list_conn_str="select * from " & db_tabl & " order by "&db_key&" asc"
	end if
	
	
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
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
    <tr>
      <td height="25"  colspan="8" bgcolor="#ffffff"><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><font color="#135294"><strong>模板管理</strong></font></td>
            <td width="68" align="center"><a href="admin_sys_web_template_edit.asp?edit_type=add">[添加模板]</a></td>
            <td width="68" align="center"><a href="admin_sys_web_template_list.asp?default=0">[默认模板]</a></td>
          </tr>
        </table></td>
    </tr>

<tr>
  
  
<td align="center" class="plon">名称</td>
<td width="120" align="center" class="plon">ip</td>
<td width="140" align="center" class="plon">时间</td>
<td width="40" height="25" align="center" class="plon">默认</td>
  <td height="25" colspan="3" align="center" class="plon">管理</td>
  </tr> 
<%
for list_conn_num=1 to page_size
if  list_conn.eof or list_conn.bof then exit for
%>
<tr>


<td class="plon">&nbsp;<%=list_conn("web_name")%></td><td width="120" class="plon">&nbsp;<%=list_conn("add_ip")%></td><td width="140" class="plon">&nbsp;<%=list_conn("add_time")%></td>
<td width="40" align="center" class="plon">&nbsp;
  <span class="red">
<%
if int(list_conn("default"))=1 then
   response.write("√")
else
   response.write("×")
end if
%>
</span>
&nbsp;
</td>

<td width="40" align="center" class="plon">
<a href="?page=<%=page%>&del_id=<%=list_conn(db_key)%>&type_id=<%=type_id%>" onclick="return confirm('是否确定删除?');">删除</a></td>
<td width="40" align="center" class="plon">
<a href="admin_<%=db_tabl%>_detail.asp?detail_id=<%=list_conn(db_key)%>&type_id=<%=type_id%>">预览</a></td>
<td width="40" align="center" class="plon">
<a href="admin_<%=db_tabl%>_edit.asp?page=<%=page%>&edit_type=edit&edit_id=<%=list_conn(db_key)%>&type_id=<%=type_id%>">修改</a></td>
</tr>

<%
 list_conn.movenext
 next
%> 
<tr>
  <td colspan="8" class="plon">
<!--#include file="admin_all_pagination.asp"-->  </td>
  </tr> 
</table>






<%

rdel_id=request.querystring("rdel_id")      '删除的id值


'--->>判断并删除
if rdel_id<>"" and isnumeric(rdel_id) then
set list_del_conn=server.createobject("adodb.recordset")

    list_del_check_str="select * from sys_web_template_replace  where id=" & int(rdel_id)
	list_del_conn_str="delete * from sys_web_template_replace  where id=" & int(rdel_id)
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
if type_id<>"" and isnumeric(type_id) then

set rlist_conn=server.createobject("adodb.recordset")
    rlist_conn_str="select * from sys_web_template_replace where type_id="&int(type_id)&" order by id asc"
	rlist_conn.open rlist_conn_str,connstr,1,3
'------->数据分页
%>
<br />
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
  <tr>
    <td height="25"  colspan="7" bgcolor="#ffffff">&nbsp;&nbsp;<font color="#135294"> <strong> 替换管理</strong></font></td>
  </tr>


<tr>
<td align="center" class="plon">替换通配符</td>
<td width="260" align="center" class="plon">替换代码</td>
<td width="40" align="center" class="plon">默认</td>

    
    <td height="25" colspan="3" align="center" class="plon">管理</td>
  </tr>
<%
do while not rlist_conn.eof
%>
<tr>


<td class="plon">&nbsp;<%=rlist_conn("web_replace_key")%></td><td width="260" class="plon">&nbsp;<%=rlist_conn("web_replace")%></td>
<td width="40" class="plon">&nbsp;
  <span class="red">
<%
if int(rlist_conn("default"))=0 then
   response.write("√")
else
   response.write("×")
end if
%>
</span>
&nbsp;
</td>


<td width="40" align="center" class="plon">
<a href="?page=<%=page%>&rdel_id=<%=rlist_conn(db_key)%>&type_id=<%=type_id%>" onclick="return confirm('是否确定删除?');">删除</a></td>
<td width="40" align="center" class="plon">
<a href="admin_<%=db_tabl%>_detail.asp?detail_id=<%=rlist_conn(db_key)%>&type_id=<%=type_id%>">预览</a></td>
<td width="40" align="center" class="plon">
<a href="admin_sys_web_template_replace_edit.asp?page=<%=page%>&edit_type=edit&edit_id=<%=rlist_conn(db_key)%>">修改</a></td>
</tr>
<%
 rlist_conn.movenext
 loop
    rlist_conn.close
set rlist_conn=nothing

end if
%> 

<tr>
  <td colspan="7" class="plon">&nbsp;</td>
</tr>
</table>

<%
if  del_back="true" then
call show_msg("已经成功删除...")
end if
%>

