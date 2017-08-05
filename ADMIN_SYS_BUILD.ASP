<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->
<!--#include file="fanren_include/f_funtion.asp"-->
<link href="fanren_style/style.css" rel="stylesheet" type="text/css">

<%
response.buffer=true

'=========================================
'========== 获取表单参数 =================
'=========================================
dim db_form,db_max,db_key

db_form=request.form("db_form")
db_max =request.form("db_max")
db_key =request.form("db_key")

if db_form="" or isnumeric(db_max)=false then
   response.write("<script>alert('操作有误！');history.back(1);</script>")
   response.end()
end if



redim field_(db_max)
redim field_show_(db_max)
redim field_type_(db_max)
redim field_max_(db_max)
redim field_input_type_(db_max)
redim field_input_set_(db_max)
redim field_input_limit_(db_max)
redim field_input_empey_type_(db_max)
redim field_input_check_type_(db_max)



'======>> 执行事件(1): 使用 for 循环获取全部表单数据
for i=0 to db_max
    field_(i)                  = request.form("field_"&i)                  '字段名称
    field_show_(i)             = request.form("field_show_"&i)             '字段名称
    field_type_(i)             = request.form("field_type_"&i)
    field_max_(i)              = request.form("field_max_"&i)
    field_input_type_(i)       = request.form("field_input_type_"&i)
    field_input_set_(i)        = request.form("field_input_set_"&i)
    field_input_limit_(i)      = request.form("field_input_limit_"&i)
    field_input_empey_type_(i) = request.form("field_input_empey_type_"&i)
    field_input_check_type_(i) = request.form("field_input_check_type_"&i)
next
'----------------------------------------------------



'======>> 执行事件(?): 执行生成文件，终极事件处理
set sys_build_conn=server.createobject("adodb.recordset")
	sys_build_conn_str="select * from sys_web_template where default=1"
	sys_build_conn.open sys_build_conn_str,connstr,1,1
	do while not sys_build_conn.eof
    '========================================   
	   web_template= sys_build_conn("web_template")        '读取模板内容
	   web_fname   = sys_build_conn("web_fname")           '读取模板附加名称
  '--------------------------------------
      
      
      
      


          
          
          
          
table_folder = base_path&"\web_admin\"&db_form                       '每个数据表对应一个目录
if sys_build_conn("type_id")=2 then
f_files_name =table_folder&"\..\..\web_other\"&db_form&"_"&web_fname&".css" 'css样式对应的路径
response.write("已生成文件: "&f_files_name&"<br>")
response.flush()
response.clear()
   else
   f_files_name =table_folder&"\"&db_form&"_"&web_fname&".asp"      '每个模板对应的生成路径
response.write("已生成文件: "&f_files_name&"<br>")
response.flush()
response.clear() 
end if
   '--------------------------------------
	   call web_create(table_folder,0,"folder")            '生成相应的目录
	   call web_create(f_files_name,web_template,"file")   '生成相应的文件
    '========================================
	sys_build_conn.movenext
	loop
	sys_build_conn.close
set sys_build_conn=nothing
%>