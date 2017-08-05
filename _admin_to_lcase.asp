<%@language="vbscript" codepage="65001"%>
<% response.codepage=65001%>
<% response.charset="utf-8" %>
<%
 response.buffer=true
 on error resume next '容错模式
%>


<%
'***********************************
'    定义变量
'***********************************
 dim fixs,path,db_table
    '-----------
     fixs="|.asp|"    '注意定义结构
	 fixs=lcase(fixs)
    '-----------
     path=request.form("path")
	 if path<>"" then
	    session("path")=path      '记录当前目录
	 elseif path="" and session("path")<>"" then
	    path=session("path")
	 else
	    path=server.mappath("../../")
	 end if
%>






<%
'***********************************
'    判断目录是否存在
'***********************************
function checkdir(ckdirname)
   dim m_fso
   checkdir=false
   set m_fso=createobject("scripting.filesystemobject")
       if (m_fso.folderexists(ckdirname)) then
           checkdir=true
       end if
   set m_fso = nothing
end function


'***********************************
'    用于输出提示行
'***********************************
sub print(str)
   response.write("<li onmouseover=""this.classname='on';"" onmouseout=""this.classname='out';""> "&str&"</li>")
   response.flush()
end sub


'***********************************
'    样式函数   
'***********************************
function css(str,s)
    css="<span class="""&s&""" >"&str&"</span>"
end function







function web_create(s_name,s_centent,s_type)
    on error resume next  '容错模式(防止内存不足)
    select case s_type
		   case "file"                            '文件不存在则生成文件
				set stm=server.createobject("adodb.stream") 
   				    stm.type=2 '以本模式读取 
   				    stm.mode=3 
    				stm.charset="utf-8"
    				stm.open
					stm.writetext s_centent 
    				stm.savetofile s_name,2 
    				stm.flush 
    				stm.close
				set stm=nothing

				
		   case "getfile"                        '读取文件内容
				set stm=server.createobject("adodb.stream") 
    				stm.type=2 '以本模式读取 
    				stm.mode=3 
    				stm.charset="utf-8"
    				stm.open
					stm.loadfromfile s_name
    				str=stm.readtext 
    				stm.close
				set stm=nothing 
   				    web_create=str
					
				set fs=server.createobject("scripting.filesystemobject")
                    path=server.mappath(filename)
                   if fs.fileexists(s_name) then
                      fs.deletefile s_name,true
                   end if
                set fs=nothing 
					
			
    end select
end function



'***********************************
'   查找分类目录下的产品，并录入到数据库
'***********************************
function dirproduct(topath)
  on error resume next     '容错模式
  set objfso = createobject("scripting.filesystemobject")
  set crntfolder = objfso.getfolder(topath)
      f_file=false
      for each confile in crntfolder.files
		 f_name=confile.name
		 
		 if len(f_name)>4 then
		    f_fix=lcase(right(f_name,4))
			f_fix="|"&f_fix&"|"
			
		   '*********** |检测文件类型是否符合| ************
			if instr(fixs,f_fix)<>0 then
         '######################## |写入数据库（重要） | ######################### 
               file_path=lcase(topath&"\"&confile.name)
			   file_path=replace(file_path,"\\","\")
			   
			   sss=lcase(web_create(file_path,,"getfile"))
               call web_create(file_path,sss,"file")
		 '####################################################################

			   if err=0 then    '数据库成功记录
			      print("&nbsp;- 已处理: "&confile.name&"  …… "&css("√","green"))
			   else
			      print("&nbsp;- 处理失败: "&confile.name&"  …… "&css("×","red"))
			   end if
			   f_file=true
			   
		    end if
		 end if


      next
  set crntfolder = nothing
  set objfso = nothing
  
 '#### 在指定目录下未检测到 相应内容 ####
  if f_file=false then call print("&nbsp;- 目录:"&topath&" ,未检测到指定类型 "&css(fixs,"strong")&" 文件  …… "&css("failed !","red"))
end function
%>



<!doctype html public "-//w3c//dtd xhtml 1.0 transitional//en" "http://www.w3.org/tr/xhtml1/dtd/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<style>
body{
font-size:12px;
color:#333333;
font-family:verdana, arial, helvetica, sans-serif;
line-height:20px;
}

body,tr,td,div{
font-size:11px;}

form{
margin:0;}

input{
	border:#666666 1px solid;
	border-bottom:#ffffff 1px solid;
	border-right:#ffffff 1px solid;
	background-color:#e2e2c7;
	color:#333333;
	width:100%;
	line-height: 18px;
	height: 18px;
}

a{
color:#cccccc;}


ul{
margin:12px;
padding-left:12px;
margin-top:0;
margin-bottom:0;}
li{
padding-left:6px;}

hr{
width:100%;
text-align:left;
background-color:#666666;
border-bottom:#666666 1px dotted;}

.red{
color:#ff0000;
font-size:10px;}
.green{
color:#00ff00;
font-size:10px;}
.strong{
color:#ff0000;
font-weight:bold;
font-size:12px;}

#main_div{
width:635px;
margin:auto;
background-color:#ebebd8;
padding:12px;}
#main_div a{
text-decoration:none;
font-size:11px;
color:#eaeaea;}
#main_div a:hover div{
background-color:#ebebd8;}



#main_box{
width:600px;
margin:auto;
background-color:#f4f4ea;
padding:15px;
border:#cccccc 1px dotted;
border-bottom:#cccccc 1px dotted;
border-right:#cccccc 1px dotted;
}



/*目录样式*/
#folders a{
float:left;
padding-top:15px;
padding-left:15px;
width:125px;
height:40px;
color:#fb9700;
text-decoration:none;
font-weight:bold}
#folders a:hover{
color:#ff0000;}


.on{
background-color:#ebebd8}
.out{
background-color:#f4f4ea;}
</style>


<title>企业产品信息录入程序 v1.2   time:12:56 2009-10-31   [卡米.伊凡 -  for 合优网络]</title>
<%if request.querystring("page")<>"folder" then%>
<script type="text/javascript">
<!--
function mm_openbrwindow(theurl,winname,features) { //v2.0
  window.open(theurl,winname,features);
}
//-->
</script>



<div id="main_div">
<div id="main_box">
<form id="addform" name="addform" method="post" action="">
<table width="100%" border="0" cellpadding="0" cellspacing="5" onmouseover="this.bgcolor='#ebebd8';" onmouseout="this.bgcolor='';">
  <tr>
    <td>
      <input name="path" type="text" id="path" value="<%=path%>" size="80%" />
	  <input name="p_form" type="hidden" id="p_form" value="ok"/>
	  </td>
    <td width="60" style="padding-left:5px;"><input name="submit" type="button" onclick="mm_openbrwindow('?page=folder','folder','width=380,height=280')" value="浏览"/></td>
    <td width="60" style="padding-left:5px;"><input type="submit" name="submit2" value="录入" /></td>
  </tr>
</table>
</form>



<ul type="a">
<%
dim i,tempstr,flspace
  
if path<>"" and request.form("p_form")="ok" then
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  onpath=server.mappath("../../")
  xpath=replace(lcase(path),lcase(onpath),"")   '获取与跟目录的相对目录
  xpath="../../"&xpath
  

  path_x_s=xpath
  path_x_s=replace(path_x_s,"\","/")
  path_x_s=replace(path_x_s,"//","/")
  path_s=server.mappath(path_x_s)     '小图全目录


'-------- 判断路径是否有效,是否存在 ---------
  call print(css("echo ↓","green"))
  'response.write("<br />")
  call print("当前目录:"&path)
  
     path_ok=true

'-------- 读取小图目录， 获取分类文件夹---------
if path_ok then
  '同时检测到大图和小图目录，才进行读取文件(go)
	  
  call print("正在分析目录  …… ")
  call dirproduct(path_s)
  
  set objfso = createobject("scripting.filesystemobject")
  set crntfolder = objfso.getfolder(path_s)

  for each subfolder in crntfolder.subfolders
	  f_id=subfolder.name
      call dirproduct(path_s&"\"&f_id)
  next
  
  set crntfolder = nothing
  set objfso = nothing
 '同时检测到大图和小图目录，才进行读取文件(end)
end if
  
  response.write("<br /><br />")
if path_ok then
   call print(css("完成录入!","strong")&"  …… "&css("ok !","green"))  
else
   call print(css("条件不符,录入失败!","strong")&"  …… "&css("failed !","red"))  
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
end if
  
%>

</ul>
</div>
</div>










<%else%>
<!--页面：文件夹选择-->

<%

 onpath=request.querystring("path")
 if session("onpath")="" and onpath<>"" then
	if checkdir(server.mappath("../../")&"/"&onpath) then session("onpath")=server.mappath("../../")&"\"&onpath
 elseif onpath<>"" then
	if checkdir(session("onpath")&"/"&onpath) then session("onpath")=session("onpath")&"\"&onpath
 else
	session("onpath")=server.mappath("../../")
 end if
 'session.abandon()
response.write(session("onpath"))
%>

<script>
function open_path(str){
	   window.location.href="?page=<%=request("page")%>&path="+str;  
}
function to_path(str){
   var path;
       path="<%=replace(session("onpath"),"\","\\")%>\\";
	   window.opener.addform.path.value=path+str;
	   window.topath.topath_str.value=path+str;
     //window.close();
}

function goto_path(){
    if(window.topath.topath_str.value!=""){
	  window.opener.addform.submit();
	  window.close();
	}else{
	  alert("请选择目录!");
	}
	   window.topath.topath_str.value=path+str;

}

</script>
<body style="margin:0; padding:0;">
<form name="topath">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="formy formaargin" >
  <tr>
    <td width="60%" class="forumrow" style="padding-left:10px;"><div id="on_path">
      <input name="topath_str" type="text" id="topath_str" readonly />
    </div></td>
    <td width="15%" align="center" class="forumrow"><input type="button" name="submit3" value="录入"  onclick="goto_path();"/></td>
    <td width="15%" align="center" class="forumrow" style="padding-right:5px;"><a href="?page=<%=request("page")%>&path=..\">上一级</a></td>
  </tr>
</table>
</form>
<div id="folders">
<%
'---------------------------------------
 set objfso = createobject("scripting.filesystemobject")
  set crntfolder = objfso.getfolder(session("onpath"))

  for each subfolder in crntfolder.subfolders
	  f_id=subfolder.name
%>
<a href="javascript:;" onclick='to_path("<%=f_id%>");' ondblclick='open_path("<%=f_id%>");'><font face='wingdings' size='5'>1</font><%=f_id%></a>
<%
  next

  set crntfolder = nothing
  set objfso = nothing
%>
</div>
</body>
<%end if%>

