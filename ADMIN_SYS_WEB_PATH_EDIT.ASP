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
    db_tabl = "sys_web_path"
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
          

f_path_name=cstr(request.form("f_path_name"))
f_path_note=cstr(request.form("f_path_note"))
f_default= int(request.form("f_default"))

      
'如：f_$字段$  =  request.form("f_$字段$")
         

call form_check("path_name",f_path_name,"limit_1")
call form_check("path_note",f_path_note,"limit_1")
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
if edit_conn("ttype")<>1 then
   '非默认目录，则可以编辑
   f_path_name1=edit_conn("path_name")
   edit_conn("path_name")=f_path_name
end if

edit_conn("path_note")=f_path_note
edit_conn("default")=f_default

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
if edit_type="add" then
   call write_web_log("新建目录路径："&f_path_name)   '记录目录路径
else
   call write_web_log("修改目录路径："&f_path_name1&" &gt;&gt; "&f_path_name) 
end if



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

f_path_name=read_conn("path_name")
f_path_note=read_conn("path_note")
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
      <td height="25"  colspan="2" bgcolor="#ffffff">&nbsp;&nbsp;<font color="#135294"> <strong> 新建目录</strong></font></td>
    </tr>

    <tr>
      <td width="36%" height="23" align="right" class="plon">目录：</td>
      <td width="64%" class="plan"><input name="f_path_name" type='text'  value='<%=f_path_name%>' /></td>
    </tr>
    <tr>
      <td height="23" align="right" class="plon">注释：</td>
      <td class="plan"><input name="f_path_note" type='text'  value='<%=f_path_note%>' /></td>
    </tr>
    <tr>
      <td height="23" align="right" class="plon">默认：</td>
      <td class="plan">
  <select name="f_default" id="f_default">
    <option value="0" <%if int(f_default)=0 then%>selected="selected"<%end if%> >√</option>
    <option value="1" <%if int(f_default)=1 then%>selected="selected"<%end if%>>×</option>
  </select>
      </td>
    </tr>
    <tr>
      <td height="23" class="plon">&nbsp;</td>
      <td class="plan"><input name="button" type="reset" class="edit_buttom" id="button" value="取消" />
        <input name="button" type="submit" class="edit_buttom" id="button" value="提交" />
        <input name="form_type" type="hidden" value="<%=edit_type%>" /></td>
    </tr>
  </table>
  </form>

<script>
function checkform()
{

if(document.admin_edit_form.f_path_name.value=="")
{alert('请填写 path_name ！');
document.admin_edit_form.f_path_name.focus();
return false;
}
if(document.admin_edit_form.f_path_note.value=="")
{alert('请填写 path_note ！');
document.admin_edit_form.f_path_note.focus();
return false;
}

if(document.admin_edit_form.f_default.value=="")
{alert('请填写 default ！');
document.admin_edit_form.f_default.focus();
return false;
}


}
</script>

