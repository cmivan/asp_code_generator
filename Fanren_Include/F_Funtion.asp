<%
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\/-->    名称：凡人网站生成系统           \/\/\/\
'\/-->    作者：凡人                       \/\/\/\
'\/-->    联系：619835864                  \/\/\/\
'\/-->    邮箱：619835864@qq.com           \/\/\/\
'\/-->    网站：http://www.fanr.com        \/\/\/\
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\



'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'\/\/\/\/\/\/\/\/定义变量 . 接收参数/\/\/\/\/\/\/\/
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
dim web_db,web_ctrl,web_db_type,web_title,web_now,web_path
  
   web_db      = request.form("web_db")
   web_ctrl    = request.form("web_ctrl")
   web_db_type = request.form("web_db_type")
   web_title   = request.form("web_title")
   web_now     = now()
   
   '--- 文件存放路径处理
   web_path       = "wwwroot\"     '生成新站存放目录
   base_path      = web_path&session("web_no.")
%>




<%
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'\/\/\/\/\   写入系统日志    \/\/\/\/\/\/\/\/\/\/\/
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

sub write_web_log(web_event)    
	        set write_web_edit_conn=server.createobject("adodb.recordset")
			    write_web_edit_conn_str="select * from sys_web_log"
			    write_web_edit_conn.open write_web_edit_conn_str,connstr,1,3
				write_web_edit_conn.addnew

   add_ip=request.servervariables("http_x_forwarded_for") 
if add_ip="" then add_ip=request.servervariables("remote_addr") 
				
			    write_web_edit_conn("add_ip")    = add_ip 
				write_web_edit_conn("add_event") = web_event
				write_web_edit_conn("add_no")    = session("web_no.")

			write_web_edit_conn.update
		    write_web_edit_conn.close
		set write_web_edit_conn= nothing
end sub
%>



<%
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'\/\/\/\/\   读取模板名称的函数    \/\/\/\/\/\/\/\/
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

sub get_tmpl_name(get_table,get_files,type_id,get_type)
    db_tabl = get_table                          '当前操作的表
	db_key="id"                                  '主键

    '------->读取数据
    set list_conn=server.createobject("adodb.recordset")
        list_conn_str="select * from " & db_tabl & " order by "&db_key&" asc"
    	list_conn.open list_conn_str,connstr,1,1
	do while not list_conn.eof
%>       
<%if get_type="li" then%>
<li><a target="main" href="admin_sys_web_template_edit.asp?edit_type=edit&edit_id=<%=list_conn(db_key)%>"><%=list_conn(get_files)%></a></li> 
<%else%> 
<option value="<%=list_conn(db_key)%>" <%if int(type_id)=int(list_conn(db_key)) then%>selected="selected"<%end if%> ><%=list_conn(get_files)%></option>
<%end if%> 
<%
        list_conn.movenext
        loop
 
        list_conn.close
    set list_conn=nothing
end sub
%>                               





<%
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'\/\/\/\/\   生成目录或文件 函数    \/\/\/\/\/\/\/\
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
sub web_create(s_name,s_centent,s_type)
                s_type=ucase(s_type)
    select case s_type
	
	       case "folder"                          '文件夹不存在则生成
 		        set fso=createobject("scripting.filesystemobject")
                    if fso.folderexists(server.mappath(s_name))=false then  
                       fso.createfolder(server.mappath(s_name))
                    end if                            
				set fso=nothing 
	   
		   case "file"                            '文件不存在则生成文件
           
       '-------生成方式1---------
 		        'set fso=createobject("scripting.filesystemobject")
		        'set savefile=fso.opentextfile(server.mappath(s_name),2,true)
		            'savefile.writeline(s_centent)
				'set savefile=nothing
				'set fso=nothing		
                
       '-------生成方式2--------- 
                set utf_file=server.createobject("adodb.stream")    
                    utf_file.type=2    
                    utf_file.mode=3    
                    utf_file.charset="utf-8"    
                    utf_file.open()    
                    utf_file.writetext s_centent    
                    utf_file.savetofile server.mappath(s_name),2    
                    utf_file.close() 
                set utf_file=nothing
       '-----------------------------------------------------------------------	    
                  		
    end select
end sub


'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'\/\/\/\/\   文件复制、  函数       \/\/\/\/\/\/\/\
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
sub copyfiles(tempsource,tempend) 
    dim fso 
    set fso = server.createobject("scripting.filesystemobject") 
if fso.fileexists(tempend) then 
       response.write "目标备份文件 <b>" & tempend & "</b> 已存在!" 
       set fso=nothing 
       exit sub 
    end if 
    if fso.fileexists(tempsource) then 
    else 
       response.write "要复制的源数据库文件 <b>"&tempsource&"</b> 不存在!" 
       set fso=nothing 
       exit sub
    end if 
    fso.copyfile tempsource,tempend 
    response.write "已经成功复制文件 <b>"&tempsource&"</b> 到 <b>"&tempend&"</b>" 
    set fso = nothing 
end sub
%>




<%
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'\/\/\/\/\  返回读取数据库的字段类型     \/\/\/\/\/
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
sub get_type(num)
    select case num
	       case "3"
		        response.write("数字")
		   case "135"
		        response.write("时间")
		   case "202"
		        response.write("文本")
		   case "203"
		        response.write("注备")
    end select
end sub
%>



<%
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'\/\/\/\/\   重选数据库     \/\/\/\/\/\/\/\/\/\/\/\
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
sub reget_web_db()
    if request.querystring("reget_web_db")="yes" then
	   session("web_no.")=""
	   response.cookies("web_no.")=""
	   call write_web_log("重选数据库文件")  '//记录网站操作事件-重选数据库	
	end if
end sub
call reget_web_db()
%>





<%
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\/\/\/\/   读取数据库表段的函数  \/\/\/\/\/\/\/\/\/\/
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
sub get_web_db_table(gtype)

dim sys_conns,sys_connstr,sys_mdb
	on error resume next
    sys_connstr="driver=microsoft access driver (*.mdb);dbq=" + session("web_db_fullpath")
    '------------------------ 	
set sys_conns=server.createobject("adodb.connection") 
    sys_conns.open sys_connstr

'\/\/\/\/\/\/\判断数据库文件是否存在\/\/\/\/\/\/\/\/\/\/\/
if err then
   err.clear
set sys_connstr = nothing
   response.write "<script>alert('数据库文件已不存在，请重选...');parent.location.href='admin_default.asp';</script>"
   
   session("web_no.")          = ""
   response.cookies("web_no.") = ""
   
   response.end
end if
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\

'------------------------
set objconn=server.createobject("adodb.connection")
    objconn.open sys_connstr
'------------------------
set rsschema=objconn.openschema(20)
    rsschema.movefirst
'------------------------
do until rsschema.eof
   if rsschema("table_type")="table" then
'------------------------  

sys_table     =ucase(request.querystring("sys_table"))
sys_table_name=ucase(rsschema("table_name"))
%>

<%if gtype="list" then%>

<li><%if sys_table=sys_table_name then%>
<a href='javascript:' class="li"><span>-&nbsp;</span><%=sys_table_name%></a>
<%else%><a href='admin_sys_web_template.asp?sys_table=<%=sys_table_name%>' onclick="location.href='?show_db=0&sys_table=<%=sys_table_name%>'" target="main">
<span>+&nbsp;</span><%=sys_table_name%></a><%end if%></li>

<%else%>
<option value="<%=sys_table_name%>" ><%=sys_table_name%></option>
<%end if%>

<%     
    end if
    rsschema.movenext
    loop
set objconn=nothing

end sub
%>





<%
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\/\/\/\/  读取代码收集分类的函数  \/\/\/\/\/\/\/\/\/\
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
sub get_code_type(gtype,gid)
    set list_code_type_conn=server.createobject("adodb.recordset")
        if gtype="1" then
              
           list_code_type_conn_str="select top 1 * from sys_web_code_type where id="&int(gid)
    	   list_code_type_conn.open list_code_type_conn_str,connstr,1,1
           if not list_code_type_conn.eof then
              response.write("[<a href='admin_sys_web_code_list.asp?code_type_id="&list_code_type_conn("id")&"'>"&list_code_type_conn("code_type_name")&"</a>]")
           end if
               list_code_type_conn.close
           '----------------------------------
           else
           
           list_code_type_conn_str="select * from sys_web_code_type order by id desc"
    	   list_code_type_conn.open list_code_type_conn_str,connstr,1,1
           if not list_code_type_conn.eof then
              response.write("["&list_code_type_conn("code_type_name")&"]")
           end if
               list_conn.close
         end if
         
    set list_code_type_conn = nothing              
end sub
%>







<%
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\/\/\/\/  读取读取模板分类的附属名称  \/\/\/\/\/\/\/\
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
function get_web_fname(gid)
     set get_web_fname_conn=server.createobject("adodb.recordset")
         get_web_fname_conn_str="select top 1 web_path,id from sys_web_template_name where id="&int(gid)
    	 get_web_fname_conn.open get_web_fname_conn_str,connstr,1,1
         if not get_web_fname_conn.eof then
            get_web_fname=get_web_fname_conn("web_path")
         end if
         get_web_fname_conn.close
     set get_web_fname_conn = nothing             
end function
%>






<%
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\/\/\/\/  读取读取模板分类的附属名称  \/\/\/\/\/\/\/\
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
function get_web_fname(gid)
     set get_web_fname_conn=server.createobject("adodb.recordset")
         get_web_fname_conn_str="select top 1 web_path,id from sys_web_template_name where id="&int(gid)
    	 get_web_fname_conn.open get_web_fname_conn_str,connstr,1,1
         if not get_web_fname_conn.eof then
            get_web_fname=get_web_fname_conn("web_path")
         end if
         get_web_fname_conn.close
     set get_web_fname_conn = nothing             
end function
%>












<%
'============================================================================
'-->> 返回信息框的  strtocode 函数
'-->> 字符替换，将特定的字符替换成要生成的网站代码
'============================================================================

'==>> 编辑页面
dim field_f,field_w,field_r,field_td,field_form,field_js

'==>> 列表页面
dim admin_list_db_1,admin_list_db_2

'==>> 详细页面
dim admin_detail_db_1

'==>> 数据库链接页面
dim data_conn_path



function strtocode(mystr)
	'==>> 公共的
		 '过滤字符 "&lt;" "&gt;"
          session(mystr)=replace(session(mystr),"&lt;","<")
	      session(mystr)=replace(session(mystr),"&gt;",">")
		 
	     mystr=replace(mystr,"$table_name$",table_name)
		 mystr=replace(mystr,"$table_key$",table_key)
		 mystr=replace(mystr,"$database_conn$",database_conn)
		
	
'==>> 编辑页面
	     mystr=replace(mystr,"$接收数据组$",field_f)        '编辑页面
		 mystr=replace(mystr,"$写入数据组$",field_w)        '编辑页面
		 mystr=replace(mystr,"$读取数据组$",field_r)        '编辑页面
	     mystr=replace(mystr,"$表单数据组$",field_td)       '生成表格
		 mystr=replace(mystr,"$数据判断$",field_form)       '数据判断
		 mystr=replace(mystr,"$表单判断$",field_js)         '表单判断
		 
'==>> 列表页面	
		 mystr=replace(mystr,"$admin_list_db_1$",admin_list_db_1)
		 mystr=replace(mystr,"$admin_list_db_2$",admin_list_db_2)
		
'==>> 详细页面
	     mystr=replace(mystr,"$admin_detail_db_1$",admin_detail_db_1)
	
'==>> 数据库链接页面
		 data_conn_path  ="..\data_"  & web_date & "\data_" & web_date & web_database_type	
		 mystr=replace(mystr,"$data_conn_path$",data_conn_path)  '表单判断
		 
end function




'============================================================================
'-->> 返回表格   <table> 、<tr> 、<td> 、<div>等  totable 函数
'-->> 指定的字符替换成要生成的网站 html 代码
'============================================================================

function totable(tablestr,t_type)
         totable="<"&t_type&">"&tablestr&"</"&t_type&">"
end function
%>