<!--#include file="../../VAR.asp"-->
<!--#include file="../function.inc"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>�鿴����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:����;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
-->
</style>
<script language="javascript">
<!--
function js_callpages(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=330");
  newwin.focus();
  return false;
}
function submitme(){
  js_callpages(frm.action);
  frm.submit();
}
--> 
</script>
</head>
<body bgcolor="#CCCCFF">
<%if Request.QueryString("id")="" or isnull(Request.QueryString("id")) then
     Response.Write "<div align='center'>�Բ��������������򿪱�ҳ��<br><a href='javascript:window.close();'>��ر�</a>"
     Response.End
  end if
  id=request("id")
  set conn=server.CreateObject("ADODB.Connection")
  Conn.open OpenString
  sql="select * from articles where id="&id
  set rs=conn.Execute(sql)
  if rs.eof then
     Response.Write "<div align='center'>�Բ������ݿ���ִ��󣬿���ԭ�򣺸�ƪ�����Ѿ���ɾ����<br><a href='javascript:window.close();'>��ر�</a>"
     Response.End
  end if
  name=trim(rs("author"))
  pic=trim(rs("face"))
  title=trim(rs("title"))
  story=trim(rs("content"))
  datetime1=trim(rs("ondate"))&trim(rs("ontime"))
  rs.close
  set rs=nothing%> 
<table width="100%" border="1" cellspacing="0" cellpadding="0" bgcolor="#3399FF" bordercolorlight="#000066" bordercolordark="#CCFFFF">
  <tr align="center" bgcolor="#FFFFFF"> 
   <td width="79%">
      <p><img src="../img/<%=pic%>.gif">&nbsp;<font color="#990099"><%=name%></font>��<font color="#990000"><%=datetime1%></font>�ڴ���<font color="#009900">��<%=title%>��</font>���ᵽ��</p>
      </td>
  </tr>
  <tr align="center"> 
    <td width="79%" align="left"> <%
       story=ViewImg(story)
       Response.Write "&nbsp;&nbsp;"&story
       %> </td>  
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
