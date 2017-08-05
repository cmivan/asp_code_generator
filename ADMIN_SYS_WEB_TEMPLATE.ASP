<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->
<!--#include file="fanren_include/f_funtion.asp"-->
<%
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\   提交数据处理-并生成文件      \\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
%>


<link href="fanren_style/style.css" rel="stylesheet" type="text/css">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="1">
<tr>
  <td valign="top">

<%
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\   数据库连接      \\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
dim sys_conns,sys_connstr,sys_mdb

    on error resume next
    sys_connstr="driver=microsoft access driver (*.mdb);dbq=" + session("web_db")
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
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

%>




</td>
  
<td width="100%" rowspan="2" valign="top">
<div class="filed" id="filed">

<span class="plon">
<%
   sys_table=request.querystring("sys_table")
if sys_table<>"" then

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
call write_web_log("选择数据表："&sys_table)   '记录选择表名数据

%>
</span>

<form id="form1" name="form1" method="post" action="admin_sys_build.asp">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
  <tr>
    <td height="25"  colspan="12" bgcolor="#ffffff">&nbsp;&nbsp;  <font color="#135294"><strong>数据库操作</strong></font></td>
  </tr>
  
        <tr>
          <td width="30" height="20" align="center" class="plon">选择          </td>
          <td width="100" align="center" class="plon">描述</td>
          <td align="center" class="plon">类型/字段</td>
          <td width="40" align="center" class="plon">限长</td>
          <td align="center" class="plon">输入框</td>
          <td width="60" align="center" class="plon">下拉设置</td>
          <td align="center" class="plon">输入限制</td>
          <td width="35" align="center" class="plon">非空</td>
          <td width="30" align="center" class="plon">主键</td>
          <td width="30" align="center" class="plon">js</td>
        </tr>
        
<%
set show_filed_conn = server.createobject("adodb.recordset")
    show_filed_conn_str = "select * from " & sys_table
    show_filed_conn.open show_filed_conn_str,sys_connstr,0,1
	for i=0 to show_filed_conn.fields.count-1
%>

<tr><td align="center" class="plon">
<input name="field_<%=i%>" type="checkbox" id="field_<%=i%>" value="<%=show_filed_conn.fields.item(i).name%>" checked="checked" />
</td>

<td width="100" class="plon" style="padding-left:5px;">
<input name="field_show_<%=i%>" type="text" id="field_show_<%=i%>" style="width:100px" value='<%=show_filed_conn.fields.item(i).name%>' /></td>

<td class="plon">
<span style="padding-left:5px;">
<input name="field_type_<%=i%>" type="hidden" id="field_type_<%=i%>" value="<%=show_filed_conn.fields.item(i).type%>" />
<%call get_type(cstr(show_filed_conn.fields.item(i).type))%> 
<span style="color: #ff0000">&gt;</span>
<%=show_filed_conn.fields.item(i).name%></span></td>

<td width="40" align="center" class="plon">
<input name="field_max_<%=i%>" type="text" id="field_max_<%=i%>" style="width:40px" value="<%=response.write(show_filed_conn.fields.item(i).definedsize)%>"/></td>

<td width="39" align="center" class="plon">
<select name="field_input_type_<%=i%>">
<%call get_tmpl_name("sys_web_input","web_input_name",0,0)%>
</select></td>

<td width="60" align="center" class="plon">
<input ondblclick="get_input(<%=i%>);" type="text" name="field_input_set_<%=i%>" readonly="readonly" style="width:60px; cursor:hand;" id="field_input_set_<%=i%>" /></td>

<td width="80" align="center" class="plon">
<select name="field_input_limit_<%=i%>" id="field_input_limit_<%=i%>">
<%call get_tmpl_name("sys_web_input_limit","web_input_limit_name",0,0)%>
</select></td>

<td width="35" align="center" class="plon">
<input name="field_input_empey_type_<%=i%>" type="checkbox" id="field_input_empey_type_<%=i%>" value="1" checked="checked" /></td>

<td width="30" align="center" class="plon">
<input type="radio" name="db_key" <%if i=0 then%>checked="checked"<%end if%> id="radio" value="<%=i%>" /></td>

<td width="30" align="center" class="plon">
<input name="field_input_check_type_<%=i%>" type="checkbox" id="field_input_check_type_<%=i%>" checked="checked" /></td>
</tr>

<%
	next
	
	show_filed_conn.close
set	show_filed_conn=nothing
%>   
        
        <tr>
          <td align="center" class="plon">
          <input type="hidden" name="db_max" value="<%=i%>" />
          <input type="hidden" name="db_form" value="<%=sys_table%>" />          </td>
          <td colspan="11" align="left" class="plon"><input name="button3" type="reset" id="button3" value="取消" />
            <input name="button2" type="submit" id="button2" value="提交" /></td>
          </tr>
        <tr>
          <td colspan="12" align="center" class="plon">&nbsp;</td>
          </tr>
      </table>
</form>


<%end if%>
</div>

</td>
</tr>
</table>


<script>
function get_input(id){
var num;
    num=4;     //用于判断 "下拉框" 的索引

if (document.getelementbyid("field_input_type_"+id).options.selectedindex!=num){
   //alert('下拉框设置前应正确选中输入框类型');
   document.getelementbyid("field_input_type_"+id).options.selectedindex=num;
}

   new_get_input=window.showmodaldialog("admin_sys_web_input_setting.asp","","dialogwidth:400px;dialogheight:180px;");
   
   document.getelementbyid("field_input_set_"+id).value=new_get_input;
}
</script>