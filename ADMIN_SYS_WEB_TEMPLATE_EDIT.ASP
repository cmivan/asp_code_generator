<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->
<!--#include file="fanren_include/f_funtion.asp"-->
<%
'|---------------------------------------------------------------------------------------|
'|---------------------------   必要参数,接收表单参数、数据库参数    --------------------|
'                                                                                        |
dim edit_type,edit_id,form_type,eorre_info                              '                |
                                                                        '                |
    edit_type  =  request.querystring("edit_type")                      '判断 编辑、修改 |
    edit_id    =  request.querystring("edit_id")                        '编辑所需的id值  |
	form_type  =  request.form("form_type")                             '判断是否提交表单|
	eorre_info =  "<div style='font-size:12px;'>访问出错...</div>"  '出错信息        |
	   
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	
dim db_tabl,db_key                                                     '数据库参数
    db_tabl = "sys_web_template"
	db_key  = "id"
	
'|---------------------------------------------------------------------------------------


  sub eorre_back(str)
      response.write("<script>alert('" & str & "');history.go(-1);</script>")
	  response.end()
  end sub

  sub form_check(s_name,s_str,limit)
     select case limit
  
            case "limit_1"                  '判断是否为空
			     if isempty(s_str) then
				    call eorre_back("你输入的 " & s_name & " ,不能为空！")
			     end if
				 
            case "limit_2"
			     
            case "limit_3"
			     if replace(s_str,"@","")=s_str then
                    call eorre_back("你输入的 " & s_name & " 不符合，请核对！")
				 end if
                 
            case "limit_4"
			     if not isnumeric(s_str) then
				    call eorre_back("你输入的 " & s_name & " 不是数字，请核对！")
				 end if
                
            case "limit_5"
			     if not isdate(s_str) then
				    call eorre_back("你输入的 " & s_name & " 不是有效的时间，请核对！")
				 end if
     end select
  end sub

%>


<%
'---------------------------------------------------------------------
'判断是修改，还是添加||| 操作数据部分                                |
'---------------------------------------------------------------------
if not isempty(form_type) then

'---------接收表单数据(循环多项的)======>>> 
          

f_web_name =cstr(request.form("f_web_name"))
f_web_fname=cstr(request.form("f_web_fname"))
f_web_template=cstr(request.form("f_web_template"))
f_type_id= int(request.form("f_type_id"))
f_default= int(request.form("f_default"))

'if f_type_id<>"" and isnumeric(f_type_id) then f_web_fname=get_web_fname(f_type_id)

   add_ip=request.servervariables("http_x_forwarded_for") 
if add_ip="" then add_ip=request.servervariables("remote_addr")

      
'如：f_$字段$  =  request.form("f_$字段$")

call form_check("web_name",f_web_name,"limit_1")
call form_check("web_template",f_web_template,"limit_1")
call form_check("type_id",f_type_id,"limit_1")
call form_check("default",f_default,"limit_1")

         set edit_conn=server.createobject("adodb.recordset")
         
			 if edit_type="add" then
		        edit_conn_str="select * from " & db_tabl
			    edit_conn.open edit_conn_str,connstr,1,3
				edit_conn.addnew
				back_info="=> " & now() & ",添加成功！"

		 elseif edit_type="edit" and not isempty(edit_id) and isnumeric(edit_id) then
                edit_conn_str="select * from " & db_tabl & " where " & db_key & "=" & int(edit_id)
				edit_conn.open edit_conn_str,connstr,1,3
				back_info="=> " & now() & ",修改成功！"

            else
			
			    response.write(eorre_info)
                response.end() 	
		     end if
 
 
'---------判断并写入数据======>>>
			 
			 if not edit_conn.eof then
	
edit_conn("add_ip")    = add_ip
edit_conn("web_name")  = f_web_name
edit_conn("web_fname") = f_web_fname
edit_conn("web_template")=f_web_template
edit_conn("type_id")   = f_type_id
edit_conn("default")   = f_default

				'如：edit_conn("$字段$")=f_$字段$
				response.write("<script>alert('" & back_info & "');</script>")
				back_info=""
			 end if

			edit_conn.update
		    edit_conn.close
		set edit_conn= nothing
end if

%>

<%
'-----------------------------------------------------------------------
'判断是修改，还是添加||| 显示数据读取部分
'-----------------------------------------------------------------------

    if edit_type="add" then
	
       'f_$字段$=request.form("f_$字段$")     [47行已经接收,该行可以忽略]

elseif edit_type="edit" and not isempty(edit_id) and isnumeric(edit_id) then
       set read_conn=server.createobject("adodb.recordset")
	       read_conn_str="select * from " & db_tabl & " where " & db_key & "=" & int(edit_id)
		   read_conn.open read_conn_str,connstr,1,3
		   if not read_conn.eof then

f_add_time=read_conn("add_time")
f_add_ip=read_conn("add_ip")
f_web_name=read_conn("web_name")
f_web_fname=read_conn("web_fname")
f_web_template=read_conn("web_template")
f_type_id=read_conn("type_id")
f_default=read_conn("default")

		   '如：f_$字段$=read_conn("$字段$")

		   else
	       response.write(eorre_info)
           response.end() 	   
		   end if
		   read_conn.close
	   set read_conn=nothing
else  
      response.write(eorre_info)
      response.end()   
end if
%>
<link href="fanren_style/style.css" rel="stylesheet" type="text/css">
<form id="admin_edit_form" name="admin_edit_form" method="post" action="?edit_type=<%=edit_type%>&edit_id=<%=edit_id%>" onsubmit="return checkform();" >
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
  <tr>
    <td height="25"  colspan="4" bgcolor="#ffffff">&nbsp;&nbsp;<font color="#135294"> <strong> 模板管理</strong></font></td>
  </tr><tr>
<td width="8%" align='right' class="plon">
名称：</td>
<td colspan="3" class="plon"><table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="29%"><input name="f_web_name" type='text'  value='<%=f_web_name%>' /></td>
    <td width="12%" align="right">附属名称：</td>
    <td width="12%"><input name="f_web_fname" type='text' id="f_web_fname"  value='<%=f_web_fname%>' /></td>
    <td width="22%">&nbsp;</td>
    <td width="25%">&nbsp;</td>
  </tr>
  
</table></td>
  </tr><tr>
<td align='right' valign="top" class="plon">
代码：</td>
<td colspan="3" class="plon">
  <textarea name="f_web_template" style="width:100%" rows="26"><%=f_web_template%></textarea></td></tr><tr>
<td align='right' class="plon">
分类：</td>
<td width="25%" class="plon">
<select name="f_type_id" id="f_type_id">
<%call get_tmpl_name("sys_web_template_name","web_name",f_type_id,0)%>
</select></td>
<td width="9%" align="right" class="plon">默认：</td>
  <td width="58%" class="plon">
  <select name="f_default" id="f_default">
    <option value="0" <%if int(f_default)=0 then%>selected="selected"<%end if%> >×</option>
    <option value="1" <%if int(f_default)=1 then%>selected="selected"<%end if%>>√</option>
  </select>
  </td>
  </tr>

<tr>
<td align="right" class="plon">&nbsp;</td>
<td colspan="3" class="plon">
<input name="button" type="reset" class="edit_buttom" id="button" value="取消">
<input name="button" type="submit" class="edit_buttom" id="button" value="提交">
<input name="form_type" type="hidden" value="<%=edit_type%>" /></td>
</tr>
</table>

</form>

<script>
function checkform()
{
if(document.admin_edit_form.f_web_name.value=="")
{alert('请填写 web_name ！');
document.admin_edit_form.f_web_name.focus();
return false;
}
if(document.admin_edit_form.f_web_template.value=="")
{alert('请填写 web_template ！');
document.admin_edit_form.f_web_template.focus();
return false;
}
if(document.admin_edit_form.f_type_id.value=="")
{alert('请填写 type_id ！');
document.admin_edit_form.f_type_id.focus();
return false;
}
if(document.admin_edit_form.f_default.value=="")
{alert('请填写 default ！');
document.admin_edit_form.f_default.focus();
return false;
}


}
</script>

