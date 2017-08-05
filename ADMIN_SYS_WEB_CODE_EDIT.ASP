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
    code_type_id=request.querystring("code_type_id")
	eorre_info =  "<div style='font-size:12px;'>访问出错...</div>"  '出错信息        |
	   
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	
dim db_tabl,db_key                                                     '数据库参数
    db_tabl = "sys_web_code"
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
          

f_code_type_id= int(request.form("f_code_type_id"))
f_code_name=cstr(request.form("f_code_name"))
f_code_centent=cstr(request.form("f_code_centent"))

      
'如：f_$字段$  =  request.form("f_$字段$")
         

call form_check("code_type_id",f_code_type_id,"limit_1")
call form_check("code_name",f_code_name,"limit_1")
call form_check("code_centent",f_code_centent,"limit_1")


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
	
				edit_conn("code_type_id")=f_code_type_id
edit_conn("code_name")=f_code_name
edit_conn("code_centent")=f_code_centent

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

		    f_code_type_id=read_conn("code_type_id")
f_code_name=read_conn("code_name")
f_code_centent=read_conn("code_centent")

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
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<form id="admin_edit_form" name="admin_edit_form" method="post" action="?edit_type=<%=edit_type%>&edit_id=<%=edit_id%>" onsubmit="return checkform();" >
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
  <tr>
    <td height="25" colspan="3" bgcolor="#ffffff">&nbsp;&nbsp;<font color="#135294"><strong>分类管理</strong></font></td>
  </tr>
  <tr>
      <td class="forumrow">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
        <tr>
          <td width="80" align='right' class="forumrow"> <%if f_code_type_id="" or isnumeric(f_code_type_id)=false then f_code_type_id=code_type_id%>
            名称：</td>
          <td width="89%" class="forumrow"><table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="81%"><input name="f_code_name" type='text' style='width:100%' value='<%=f_code_name%>' /></td>
              <td width="19%">
&nbsp;
<%call get_code_type(1,f_code_type_id)%>
                <input name=f_code_type_id type='hidden' value='<%=f_code_type_id%>' /></td>
              </tr>
            
          </table></td>
        </tr>
        <tr>
          <td width="80" align='right' class="forumrow"> 代码：</td>
          <td class="forumrow"><textarea name="f_code_centent" rows="12" style="width:100%"><%=f_code_centent%></textarea>            </td>
        </tr>
          <tr>
            <td width="80" align="right" class="forumrow">&nbsp;</td>
          <td class="forumrow"><input name="button" type="reset" class="edit_buttom" id="button" value="取消" />
                <input name="button" type="submit" class="edit_buttom" id="button" value="提交" />
                <input name="form_type" type="hidden" value="<%=edit_type%>" />            </td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>

<script>
function checkform()
{

if(document.admin_edit_form.f_code_type_id.value=="")
{alert('请填写 code_type_id ！');
document.admin_edit_form.f_code_type_id.focus();
return false;
}
if(document.admin_edit_form.f_code_name.value=="")
{alert('请填写 code_name ！');
document.admin_edit_form.f_code_name.focus();
return false;
}
if(document.admin_edit_form.f_code_centent.value=="")
{alert('请填写 code_centent ！');
document.admin_edit_form.f_code_centent.focus();
return false;
}


}
</script>

