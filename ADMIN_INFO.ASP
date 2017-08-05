<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->
<!--#include file="fanren_include/f_funtion.asp"-->
<%
	dim theinstalledobjects(17)
    theinstalledobjects(0) = "mswc.adrotator"
    theinstalledobjects(1) = "mswc.browsertype"
    theinstalledobjects(2) = "mswc.nextlink"
    theinstalledobjects(3) = "mswc.tools"
    theinstalledobjects(4) = "mswc.status"
    theinstalledobjects(5) = "mswc.counters"
    theinstalledobjects(6) = "iissample.contentrotator"
    theinstalledobjects(7) = "iissample.pagecounter"
    theinstalledobjects(8) = "mswc.permissionchecker"
    theinstalledobjects(9) = fs
    theinstalledobjects(10) = "adodb.connection"
    theinstalledobjects(11) = "softartisans.fileup"
    theinstalledobjects(12) = "softartisans.filemanager"
    theinstalledobjects(13) = "jmail.smtpmail"
    theinstalledobjects(14) = "cdonts.newmail"
    theinstalledobjects(15) = "persits.mailsender"
    theinstalledobjects(16) = "lyfupload.uploadfile"
    theinstalledobjects(17) = "persits.upload.1"
%>


<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" >
<link href="fanren_style/style.css" rel="stylesheet" type="text/css">
<title>后台管理首页</title>
</head>
<body style="margin-top: 0px">
<%
'判断是否已经记录cookies("web_no.") 若存在 则直接读取
if request.cookies("web_no.")<>"" and session("web_no.")="" then
   session("web_no.")=request.cookies("web_no.")
   response.write("<script>parent.location.href='admin_default.asp';</script>")
end if
%>

<%if session("web_no.")="" then%>

<%
'----------------------------
if web_db<>"" then
   web_now=replace(web_now,"-","")
   web_now=replace(web_now,":","")
   web_now=replace(web_now," ","")
   
'============================比较重要的session值记录=========================================

	   set edit_conn=server.createobject("adodb.recordset")
			   
			    edit_conn_str="select * from sys_web_date"
			    edit_conn.open edit_conn_str,connstr,1,3
				edit_conn.addnew

   add_ip=request.servervariables("http_x_forwarded_for") 
if add_ip="" then add_ip=request.servervariables("remote_addr") 
				
   session("web_db")     =web_db           '//记录网站生成的 数据库路径
   session("web_ctrl")   =web_ctrl         '//记录网站生成的 控件加密
   session("web_db_type")=web_db_type      '//记录网站生成的 数据库后缀
   session("web_title")  =web_title        '//记录网站生成的 网站名称
   session("web_no.")    =web_now          '//记录网站生成的 编号
   
   response.cookies("web_no.")        = web_now         '//使用cookies 记录编号
   response.cookies("web_no.").expires = date()+10000    '//cookies 有效时间
		
				edit_conn("add_ip")           = add_ip 
				edit_conn("web_data")         = session("web_db")
				edit_conn("web_ctrl")         = session("web_ctrl")
				edit_conn("web_data_type")    = session("web_db_type")
				edit_conn("web_name")         = session("web_title")
				edit_conn("web_no")           = session("web_no.")
        
				 response.write("<script>alert('');parent.location.href='admin_default.asp';</script>")

			edit_conn.update
		    edit_conn.close
		set edit_conn= nothing
		
'============================写入系统日志=========================================  	
call write_web_log("新建工程："&web_title&"<br>编号："&session("web_no.")&"<br>数据库路径："&web_db)  '//记录网站操作事件	


end if
%>
<br>


<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
<form action="" method="post" style="margin:0">
  <tr>
    <td height="25" colspan="3" bgcolor="#ffffff">&nbsp;&nbsp; <font color="#135294"><strong>请选择数据文件</strong></font></td>
  </tr>
  <tr>
<td width="17%" height="20" align="center" valign="middle" class="forumrow">选数据库：</td>
<td width="25" height="20" colspan="2" class="forumrow">
<input name="web_db" type="file" id="web_db" /></td>
  </tr>
  <tr>
    <td height="20" align="center" valign="middle" class="forumrow">控件加密：</td>
    <td height="20" class="forumrow"><input name="web_ctrl" type="text" id="web_ctrl" value="web_" size="20"  width="200px"/>

<select name="web_db_type">
<option value="mdb" >mdb</option>
<option value="asa" >asa</option> 
<option value="asp" >asp</option>   
</select>    </td>
    <td class="forumrow">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" align="center" valign="middle" class="forumrow">网站名称：</td>
    <td height="20" colspan="2" valign="middle" class="forumrow"><input name="web_title" type="text" id="web_title" value="凡人生成系统" size="20"  width="200px"/>      <input type="submit" name="button" id="button" value="提交" /></td>
    </tr>
  <tr>
    <td height="20" colspan="3" align="center" valign="middle" class="forumrow">
    <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
      <tr>
        <td colspan="3" class="forumraw">&nbsp;&nbsp;历史记录</td>
        </tr>
<%
    db_tabl = "sys_web_date"                  '当前操作的表
	page_size=8                               '每页显示数目

sub show_msg(str)
    response.write("<script>alert('"&str&"');</script>")
end sub

	

'------->读取数据
set list_conn=server.createobject("adodb.recordset")
    list_conn_str="select * from " & db_tabl & " order by id desc"
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
<%
for list_conn_num=1 to page_size
if  list_conn.eof or list_conn.bof then exit for
%>
      <tr>
        <td class="forumrow">·<a href="?web_no=<%=list_conn("web_no")%>"><%=list_conn("web_name")%></a></td>
        <td class="forumrow"><%=list_conn("web_no")%></td>
        <td class="forumrow"><%=list_conn("add_time")%></td>
      </tr>

<%
 list_conn.movenext
 next
    list_conn.close
set list_conn=nothing
%>
      <tr>
        <td colspan="3" class="forumrow">&nbsp;</td>
        </tr>
    </table>
    </td>
  </tr>
  </form>
</table>


<%else%>

<br />
<%
'|---------------------------------------------------------------------------------------
'|---------------------------   必要参数,接收表单参数、数据库参数    --------------------|
'                                                                                        |
dim db_tabl,db_key                            '数据库参数
dim del_id,del_back,page_size,list_num,list_conn_num        '删除返回的数据


'以上的是配置信息 ————————
    db_tabl = "sys_web_log"                          '当前操作的表
	db_key="id"                           '主键
	page_size=15                              '每页显示数目

	
sub show_msg(str)
    response.write("<script>alert('"&str&"');</script>")
end sub

	

'------->读取数据
set list_conn=server.createobject("adodb.recordset")
    list_conn_str="select * from " & db_tabl & " where add_no='"&session("web_no.")&"' order by "&db_key&" desc"
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

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
  <tr>
    <td height="25" colspan="4" bgcolor="#ffffff">&nbsp;&nbsp; <font color="#135294"><strong>系统操作日志，名称：<%=session("web_note")%> &nbsp;||&nbsp;&nbsp;编号（<%=session("web_no.")%>）</strong></font></td>
  </tr>
  <tr>
    <td width="30" height="20" align="center" valign="middle" class="forumrow">id</td>
    <td width="120" height="20" align="center" class="forumrow">time</td>
    <td width="120" align="center" class="forumrow">ip</td>
    <td align="center" class="forumrow">event</td>
  </tr>
<%
for list_conn_num=1 to page_size
if  list_conn.eof or list_conn.bof then exit for
%>
  <tr>
<td width="30" height="20" align="center" valign="middle" class="forumrow"><%=list_conn("id")%></td>
<td height="20" class="forumrow"><%=list_conn("add_time")%></td>
  <td width="120" class="forumrow"><%=list_conn("add_ip")%></td>
  <td class="forumrow"><%=list_conn("add_event")%></td>
  </tr>
<%
 list_conn.movenext
 next
    list_conn.close
set list_conn=nothing
%> 
  <tr>
    <td height="20" colspan="4" align="left" valign="middle" class="forumrow">&nbsp;&nbsp;&gt;&gt;&nbsp;当前路径：<%=server.mappath("wwwroot\")%><br>

&nbsp;&nbsp;&gt;&gt;&nbsp;生成路径：<%=server.mappath("wwwroot\"&session("web_no.")&"\")%>    </td>
  </tr>
  <tr>
    <td height="20" colspan="4" align="left" valign="middle" class="forumrow">
<!--#include file="admin_all_pagination.asp"-->
    </td>
  </tr>
</table>



<br>
<%
    db_tabl  = "sys_web_date"                            '当前操作的表
	db_key   = "id"                                      '主键
	page_size=10                                         '每页显示数目

'------->读取数据
set list_conn=server.createobject("adodb.recordset")
    list_conn_str="select * from " & db_tabl & " where web_no='"&session("web_no.")&"'"
	list_conn.open list_conn_str,connstr,1,3
'------->数据分页

        if not list_conn.eof then
   session("web_db")     =list_conn("web_data")        '//记录网站生成的 数据库路径
   session("web_ctrl")   =list_conn("web_ctrl")        '//记录网站生成的 控件加密
   session("web_db_type")=list_conn("web_data_type")   '//记录网站生成的 数据库后缀
   session("web_title")  =list_conn("web_name")        '//记录网站生成的 网站名称
   session("web_no.")    =list_conn("web_no")          '//记录网站生成的 编号
        end if
		
    list_conn.close
set list_conn=nothing

%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
<form action="" method="post" style="margin:0">
  <tr>
    <td height="25" colspan="3" bgcolor="#ffffff">&nbsp;&nbsp; <font color="#135294"><strong>数据库配置</strong></font></td>
  </tr>
  <tr>
<td width="17%" height="20" align="center" valign="middle" class="forumrow">选数据库：</td>
<td width="25" height="20" colspan="2" class="forumrow"><input name="web_db" type="text" id="web_db" value="<%=session("web_db")%>" disabled /></td>
  </tr>
  <tr>
    <td height="20" align="center" valign="middle" class="forumrow">控件加密：</td>
    <td height="20" class="forumrow">
<input disabled name="web_ctrl" type="text" id="web_ctrl" value="<%=session("web_ctrl")%>" size="20"  width="200px"/>

<select name="web_db_type" disabled>
  <option value="<%=session("web_db_type")%>"><%=session("web_db_type")%></option>   
</select>    </td>
    <td class="forumrow">&nbsp;</td>
  </tr>
  <tr>
    <td height="20" align="center" valign="middle" class="forumrow">网站名称：</td>
    <td height="20" colspan="2" valign="middle" class="forumrow"><input disabled name="web_title" type="text" id="web_title" value="<%=session("web_title")%>" size="20"  width="200px"/></td>
  </tr>
  <tr>
    <td height="20" colspan="3" align="left" valign="middle" class="forumrow">
<%

'/////////////////////////////////////////////////////////////////////////////
'////////////////// 判断是否需要生成新的网站系统目录 /////////////////////////
'/////////////////////////////////////////////////////////////////////////////
if session("web_no.")<>"" then

'先生成基本目录==============================:
    base_path            =web_path&session("web_no.")
	session("web_db_path")="\data#ivan_"&session("web_no.")&"."&session("web_db_type")

    call web_create(web_path,0,"folder")
    call web_create(base_path,0,"folder")
	
	'call web_create(base_path&"\web_data\",0,"folder")
	'生成数据库存放路径
	'call web_create(base_path&web_admin_path,0,"folder")
	'生成后台管理路径

	
	'复制数据库文件到指定的数据库目录
	   path1=session("web_db")
	   '参数1
	   path2=server.mappath(base_path&"\web_data\")&session("web_db_path")

       session("web_db_fullpath")=path2       '///重点啊，用于后边
	   '参数2

	   call copyfiles(path1,path2)
	   
	   
	   
'开始生成其他目录
'===================================

	db_tabl  = "sys_web_path"                            '当前操作的表
	db_key   = "id"                                      '主键
'------->读取数据
set list_conn=server.createobject("adodb.recordset")
    list_conn_str="select * from " & db_tabl & " where default=0"
	list_conn.open list_conn_str,connstr,1,3
    do while not list_conn.eof
	   
	   path_name=list_conn("path_name")
	   path_name=base_path&"\"&path_name
			
	   call web_create(path_name,s_centent,"folder")
	   '//创建目录 web_create("相对目录路径",s_centent,"folder")
	list_conn.movenext
    loop
		
    list_conn.close
set list_conn=nothing

end if
%>
    &nbsp;</td>
  </tr>
  </form>
</table>
<%end if%>
<table width="100%" border="0" align=center cellpadding=3 cellspacing=1 bordercolor="#ffffff" bgcolor="#c4d8ed">
      <tr>
      <td  colspan=2 bgcolor="#ffffff">&nbsp;&nbsp;<font color="#135294"> <strong> 网站信息</strong></font></td>
    </tr>
 
  <tr>
    <td  height=23 colspan="2"  class="plon">&nbsp;&nbsp;服务器类型：
	<%=request.servervariables("os")%>(ip:<%=request.servervariables("local_addr")%>)</td>　  </tr>
  <tr>
    <td width="44%" height=23 class="plon">&nbsp;&nbsp;数据库使用：
      <%if  isobjinstalled(theinstalledobjects(10)) then
	  response.write "<font color=green><b>√</b></font>"
      else
	  response.write "<font color=red><b>×</b></font>"
      end if
      %></td>
    <td width="56%" class="plan">&nbsp;&nbsp;无组件上传支持：
      <%
      if isobjinstalled("adodb.stream") then
	  response.write "<font color=green><b>√</b></font>"
      else
	  response.write "<font color=red><b>×</b></font>"
      end if
      %>    </td>
  </tr>
  <tr>
    <td height=23 class="plon">&nbsp;&nbsp;jmail组件支持：
      <%if isobjinstalled(theinstalledobjects(13)) then
	   response.write "<font color=green><b>√</b></font>"
       else
	   response.write "<font color=red><b>×</b></font>"
       end if%>    </td>
    <td class="plan">&nbsp;&nbsp;cdonts组件支持：
      <%if isobjinstalled(theinstalledobjects(14)) then
	  response.write "<font color=green><b>√</b></font>"
      else
	  response.write "<font color=red><b>×</b></font>"
      end if
      %></td>
  </tr>
  <tr>
    <td height=23 class="plon">&nbsp;&nbsp;aspjpeg组件支持：
      <%if not isobjinstalled("persits.jpeg") then
        response.write "<font color=red><b>×</b></font>"
       else
	   response.write "<font color=green><b>√</b></font>"
       end if%>    </td>
    <td class="plan">&nbsp;&nbsp;aspemail组件支持：
      <%if not isobjinstalled("persits.mailsender") then
        response.write "<font color=red><b>×</b></font>"
       else
	   response.write "<font color=green><b>√</b></font>"
       end if%></td>
  </tr>
  
  <tr>
    <td height=25 colspan=2 class="plen">&nbsp;&nbsp;请注意做好定期数据备份，数据的定期备份可最大限度的保障您网站数据的安全 &nbsp;&nbsp; </td>
  </tr>
  <tr>
    <td height=25 colspan=2 align="right" class="plen">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>　</td>
  </tr>
</table>
<br>
</body></html>
<% 
function isobjinstalled(strclassstring)
	on error resume next
	isobjinstalled = false
	err = 0
	dim xtestobj
	set xtestobj = server.createobject(strclassstring)
	if 0 = err then isobjinstalled = true
	set xtestobj = nothing
	err = 0
end function

conn.close
set conn=nothing
%>

<%
'/////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////
   web_no=request.querystring("web_no")
if web_no<>"" then 
   session("web_no.")=web_no
   response.write("<script>window.parent.location.href='admin_default.asp';</script>")
   response.end()
end if
'/////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////
%>