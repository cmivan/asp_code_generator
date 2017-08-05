<!--#include file="admin_chk.asp"-->
<!--#include file="admin_conn.asp"-->
<!--#include file="fanren_include/f_funtion.asp"-->
<link href="fanren_style/style.css" rel="stylesheet" type="text/css">

<body onunload="back_value();" style="margin:0">
<table width="400" height="100%" border="0" align="center" cellpadding="3" cellspacing="1" bordercolor="#ffffff" bgcolor="#c4d8ed">
<form id="form1" name="form1" method="post" action="">

<tr>
    <td height="25" colspan="2" bgcolor="#ffffff">&nbsp;&nbsp; <font color="#135294"><strong>请选择表段、字段</strong></font></td>
  </tr>
  <tr>
    <td width="20%" align="right" class="forumrow">数据&nbsp;.表： </td>
    <td width="80%" class="forumrow">
<select name="files_" id="files_" onchange="back_files(this.value);">
<%call get_web_db_table("0") '///读取数据库表段%>
</select>   </td>
  </tr>
  <tr>
    <td align="right" class="forumrow">字段[值]： </td>
    <td class="forumrow">
<select name="files_value" id="files_value">
</select></td>
  </tr>
  <tr>
    <td align="right" class="forumrow">字段[显]：</td>
    <td class="forumrow">
<select name="files_show" id="files_show">
</select></td>
  </tr>
  <tr>
    <td align="right" class="forumrow">显示数量：</td>
    <td class="forumrow"><label>
      <select name="files_num" id="files_num">
        <option value="*">*  全</option>
        <option value="5">5  条</option>
        <option value="8">8  条</option>
        <option value="10">10条</option>
        <option value="12">12条</option>
        <option value="15">15条</option>
        <option value="20">20条</option>
                  </select>
    </label></td>
  </tr>
  <tr>
    <td align="right" class="forumrow">&nbsp;</td>
    <td class="forumrow"><input type="button" onclick="window.close();" name="button" id="button" value="完成并关闭" /></td>
  </tr>
   </form>
</table>

<script>
//用于返回 选择数据表后返回数据字段
function back_files(tables)
{

  var source;
  var sourcename = "admin_sys_web_input_setting_xml.asp?tables_id="+tables;
  var source = new activexobject('microsoft.xmldom');   //创建一个 msxml解析器实例
  source.async = false;
  source.load(sourcename);   //装入xml文档
  root = source.documentelement;   //设置文档元素为根节点元素
  choosecity();
}  

function choosecity()
{
  sortfield=root.selectnodes("//files[@name='"+1+"']/x");   
  //搜索name属性值等于参数y的state节点下的所有city节点
  
  for(var i=form1.files_value.options.length-1;i>=0;--i)   //撤消原来的列表项
  {
    form1.files_value.options.remove(i);
	form1.files_show.options.remove(i);
  }
  for(var i=0;i<sortfield.length;++i)   //增加城市名称到下拉列表
  {
    var ooption = document.createelement('option');
        ooption.text = " "+sortfield[i].text+" ";
        ooption.value = sortfield[i].text;
	
    var ooption1 = document.createelement('option');
        ooption1.text = " "+sortfield[i].text+" ";
        ooption1.value = sortfield[i].text;

	
    form1.files_value.options.add(ooption); 
	form1.files_show.options.add(ooption1);
  }
}  

//========================================================

function back_value(){
files_   =form1.files_.value;
files_v  =form1.files_value.value;
files_s  =form1.files_show.value;
files_n  =form1.files_num.value;

if (files_v==''||files_s.value==''){
    alert('你没选择完整的数据.');
    window.returnvalue='';
} else {
    window.returnvalue=files_+'|'+files_v+'|'+files_s+'|'+files_n;
}


}

</script>

</body>