<%@language="vbscript" codepage="65001"%>
<% response.codepage=65001%>
<% response.charset="utf-8" %>
<%
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'\/-->    名称：凡人网站生成系统           \/\/\/\
'\/-->    作者：凡人                       \/\/\/\
'\/-->    联系：619835864                  \/\/\/\
'\/-->    邮箱：619835864@qq.com           \/\/\/\
'\/-->    网站：http://www.fanr.com        \/\/\/\
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\


'设置session超时时间
   session.timeout=40

'判断是否登录，是否有管理权限
if session("admin_id")="" or isnumeric(session("admin_id"))=false or session("admin_name")="" then

   ip=request.servervariables("http_x_forwarded_for")
if ip="" then ip= request.servervariables("remote_addr")
   
%>
<script>
alert("\n1.-登录超时，或权限不足！\n\n\n2.-你的ip为[<%=ip%>]，已记录...\n\n\n3.如有问题，可与管理员联系\n\n\n4.技术支持[ 新格构 www.sznt.net ]\n\n\n                        联系qq:619835864");
</script>
<%
response.end()
end if
%>