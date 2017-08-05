<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->

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
    db_tabl = "sys_web_log"
	db_key  = "id"
	
'|---------------------------------------------------------------------------------------


  sub eorre_back(str)
      response.write("<script>alert('" & str & "');history.go(-1);</script>")
	  response.end()
  end sub

  sub form_check(s_name,s_str,limit)
     select case limit
  
            case "no"                  '判断是否为空
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
          

f_add_time=request.form("f_add_time")
f_add_ip=cstr(request.form("f_add_ip"))
f_add_event=cstr(request.form("f_add_event"))

      
'如：f_$字段$  =  request.form("f_$字段$")
         

call form_check("add_time",f_add_time,"limit_1")
call form_check("add_ip",f_add_ip,"limit_1")
call form_check("add_event",f_add_event,"limit_1")


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
	
				edit_conn("add_time")=f_add_time
edit_conn("add_ip")=f_add_ip
edit_conn("add_event")=f_add_event

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
f_add_event=read_conn("add_event")

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
<link href="../stype/admin_css.css" rel="stylesheet" type="text/css" />
<form id="admin_edit_form" name="admin_edit_form" method="post" action="?edit_type=<%=edit_type%>&edit_id=<%=edit_id%>" onsubmit="return checkform();" >

<table width="100%" border="0" cellpadding="0" cellspacing="10" bgcolor="#f7f7f7">
<tr>
<td>

<table width="100%" border="0" cellpadding="0" cellspacing="2" class="edit_table">

<tr><td align='right'>
add_time：</td>
<td>
<input name=f_add_time type='text' style='width:100%' value='<%=f_add_time%>' />
</td></tr><tr><td align='right'>
add_ip：</td>
<td>
<input name=f_add_ip type='text' style='width:100%' value='<%=f_add_ip%>' />
</td></tr><tr><td align='right'>
dd_event：</td>
<td>
<input name=f_add_event type='text' style='width:100%' value='<%=f_add_event%>' />
</td></tr>

<tr>
<td align="right">&nbsp;</td>
<td>
<input name="button" type="reset" class="edit_buttom" id="button" value="取消">
<input name="button" type="submit" class="edit_buttom" id="button" value="提交">
<input name="form_type" type="hidden" value="<%=edit_type%>" />
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>

<script>
function checkform()
{

if(document.admin_edit_form.f_add_time.value=="")
{alert('请填写 add_time ！');
document.admin_edit_form.f_add_time.focus();
return false;
}
if(document.admin_edit_form.f_add_ip.value=="")
{alert('请填写 add_ip ！');
document.admin_edit_form.f_add_ip.focus();
return false;
}
if(document.admin_edit_form.f_add_event.value=="")
{alert('请填写 dd_event ！');
document.admin_edit_form.f_add_event.focus();
return false;
}


}
</script>

