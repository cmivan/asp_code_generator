<%
'所需变量 ==>>由导入的列表页面提供
'page=0               '当前页数
'list_num=3           '
'frist_page=0         ' 首页
'last_page=10         ' 最后页


num1=page-list_num
num2=page+list_num
type_id=request.querystring("type_id")
default=request.querystring("default")

if num1<1 then num1=1
if num2>last_page then num2=last_page
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center">
<div class="pagination">
当前：(<%=page%>/<%=last_page%>)&nbsp;&nbsp;&nbsp;
<a href="?type_id=<%=type_id%>&default=<%=default%>&page=1&code_type_id=<%=code_type_id%>">&lt;&lt;</a>
<%

for num=num1 to num2
%>
<%if num=page then%>
<a href="?type_id=<%=type_id%>&default=<%=default%>&page=<%=num%>&code_type_id=<%=code_type_id%>" style="color:#ff0000"><%=num%></a>
<%else%>
<a href="?type_id=<%=type_id%>&default=<%=default%>&page=<%=num%>&code_type_id=<%=code_type_id%>"><%=num%></a>
<%end if%>
<%
next
%>


<a href="?type_id=<%=type_id%>&default=<%=default%>&page=<%=last_page%>&code_type_id=<%=code_type_id%>">&gt;&gt;[<%=last_page%>]</a></div>    </td>
  </tr>
</table>

