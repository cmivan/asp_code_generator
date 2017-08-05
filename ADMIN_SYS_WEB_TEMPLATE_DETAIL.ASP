<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->

<%
'|---------------------------------------------------------------------------------------
'|---------------------------   必要参数,接收表单参数、数据库参数    --------------------|
'                                                                                        |
dim db_tabl,db_key                            '数据库参数
dim del_id,del_back,page_size,detail_num        '删除返回的数据


'以上的是配置信息 ————————
    db_tabl = "sys_web_template"           '当前操作的表
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
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<table width="100%" border="0" align=center cellpadding=3 cellspacing=1 bordercolor="#ffffff" bgcolor="#c4d8ed">
      <tr>
      <td bgcolor="#ffffff">&nbsp;&nbsp;<font color="#135294"> <strong> 模板预览</strong></font></td>
    </tr>
      <tr>
        <td class="plon">&nbsp;
<pre>
<%=server.htmlencode(detail_conn("web_template"))%>
</pre>        </td>
      </tr>
      <tr>
        <td class="plon">
<%=detail_conn("web_template")%>
        </td>
      </tr>
</table>


<br />
<%
	detail_conn.close
set detail_conn=nothing
%>

