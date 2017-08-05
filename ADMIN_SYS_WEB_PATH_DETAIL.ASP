<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->

<%
'|---------------------------------------------------------------------------------------
'|---------------------------   必要参数,接收表单参数、数据库参数    --------------------|
'                                                                                        |
dim db_tabl,db_key                            '数据库参数
dim del_id,del_back,page_size,detail_num        '删除返回的数据


'以上的是配置信息 ————————
    db_tabl = "sys_web_path"           '当前操作的表
	db_key="id"                '主键
    detail_id=request.querystring("detail_id")

'读取数据
if detail_id="" or isnumeric(detail_id)=false then
   response.write("<script>alert('参数有误...');history.go(-1);</script>")
   response.end()
end if

set detail_conn=server.createobject("adodb.recordset")
    detail_conn_str="select * from " & db_tabl & "  where "&db_key&"=" & int(detail_id)
	detail_conn.open detail_conn_str,connstr,1,3    
	if detail_conn.eof then
	   response.write("<script>alert('参数有误...');history.go(-1);</script>")
	   response.end()
	end if
%>
<link href="fanren_style/style.css" rel="stylesheet" type="text/css">


<table width="100%" border="0" align=center cellpadding=3 cellspacing=1 bordercolor="#ffffff" bgcolor="#c4d8ed">
      <tr>
      <td height="25"  colspan=2 bgcolor="#ffffff">&nbsp;&nbsp;<strong><font color="#135294">目录配置</font></strong></td>
    </tr>

    

<tr><td width="16%" class="plon">path_name</td>
<td class="plon">&nbsp;<%=detail_conn("path_name")%></td>
</tr><tr><td class="plon">path_note</td>
  <td class="plon">&nbsp;<%=detail_conn("path_note")%></td></tr><tr><td class="plon">add_time</td>
    <td class="plon">&nbsp;<%=detail_conn("add_time")%></td></tr><tr><td class="plon">default</td>
        <td class="plon">&nbsp;<%=detail_conn("default")%></td></tr>
</table>

<%
	detail_conn.close
set detail_conn=nothing
%>