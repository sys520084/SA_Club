<!--#include file="..\..\var.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><!-- InstanceBegin template="/Templates/bs.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>�ϴ��ļ�</title>
<!-- InstanceEndEditable --> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!-- InstanceBeginEditable name="head" -->
<script language="javascript">
<!--
function js_callpages(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=600,height=430");
  newwin.focus();
  return false;
}
-->
</script>
<!-- InstanceEndEditable --> 
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:white;}
a:link {text-decoration: none; color: yellow }
a:visited {
	text-decoration: none;
	color: #FFFF00;
}
a:hover {
	text-decoration: underline;
	color: #FF9900;
}
a:active {text-decoration;
	color: #FF9900;
}
.unnamed1 {  font-family: "����"; font-size: 9pt; font-style: normal; line-height: 100%; font-weight: normal; color: #99FFFF}
.unnamed2 {  font-family: "����"; font-size: 9pt; line-height: 100%; font-weight: normal; color: #294C39}
input {
	font-family: "����";
	font-size: 10pt;
	color: #FFFFFF;
	background-color: #000033;
}
textarea {
	font-family: "����";
	font-size: 9pt;
	color: #FFFFFF;
	background-color: #003366;
}
-->
</style>
</head>

<body bgcolor="#043F80" link="#00FFFF" alink="#00FFFF" leftmargin="0" topmargin="0">
<p align="center"><img src="../img/index.jpg" width="600" height="140"></p>
<!-- InstanceBeginEditable name="content" -->
<div align=center> 
  <%
  if session("user")="" then
         Response.Write "<div align='center'>�Բ���������ʹ���ϴ�������<a href='../membership/relogin.asp' onclick='return js_callpage(this.href)'>���µ�½</a>��</div>"
         Response.End
  end if
set conn=server.CreateObject("ADODB.Connection")
conn.Open OpenString
set rs=server.CreateObject("ADODB.Recordset")
sql="select * from files where filename='"&replace(request("filename"),"'","''")&"'"
'response.write sql
rs.Open sql,conn,1,3
if rs.EOF then
	Response.Write "<div align=center>�Բ������ݿ��д��󣬲��ܱ����ļ�����</div>"
	Response.End
end if



rs("title")=request("title")
rs("des")=request("des")
rs("keywords")=request("keywords")
rs("author")=request("author")
rs("journal")=request("journal")
if request("publishtime")<>"" then
rs("publishtime")=year(request("publishtime"))&"-"&month(request("publishtime"))&"-"&day(request("publishtime"))
end if
'response.end 
rs.Update
rs.Close 
sql="select id from files where filename='"&request("filename")&"'"
rs.Open sql,conn,3
id=rs("id")
rs.Close 
%>
  <p align=center>�ļ�����ɹ�</p>
  <p>
    <%
uptype=request("uptype")

if isnull(uptype) or uptype="" then
	uptype=0
end if
'response.end
	if uptype<>1 then%>
    <script language="javascript">window.location.replace("first.asp")</script>
    <%
		'response.end 

	else
			'response.end 

	%>
<script language="javascript">
window.opener.document.forms[0].content.value="<a href='../upload/download.asp?id=<%=id%>'><%=request("title")%></a>";
</script>
    ����������д��뿽������Ҫ����������У�</p>
  <p><font color="red">[upload]../upload/download.asp?id=<%=id%>[/mid]<%=request("title")%>[/upload]</font> 
    <%
	'response.end
    end if%>
  </p>
  <table width="100%" border="1" cellpadding="1" cellspacing="1" bordercolor="#000033" bgcolor="#6666FF">
    <tr> 
      <td width="36%" bgcolor="#000033">�ļ�����</td>
      <td width="64%"><%=request("filename")%></td>
    </tr>
    <tr> 
      <td bgcolor="#000033">��չ����</td>
      <td><%=request("extname")%></td>
    </tr>
    <tr> 
      <td bgcolor="#000033">�ϴ��û���</td>
      <td><%=request("user")%></td>
    </tr>
    <tr> 
      <td bgcolor="#000033">�ļ���С��</td>
      <td><%=request("bytes")%>
        <input name="bytes" type="hidden" value=<%=filesize%> id="bytes"></td>
    </tr>
    <tr> 
      <td bgcolor="#000033">�ļ����⣺</td>
      <td><%=request("title")%></td>
    </tr>
    <tr> 
      <td bgcolor="#000033">�ļ�����/����ժҪ��</td>
      <td><%=request("des")%></td>
    </tr>
		<tr> 
          <td bgcolor="#000033">�ؼ��֣�</td>
          <td><%=request("keywords")%></td>
        </tr>
		<tr> 
		<tr> 
          <td bgcolor="#000033">�������ߣ�</td>
         <td><%=request("author")%></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">������/�ļ���</td>
          <td><%=request("journal")%></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">����ʱ�䣺</td>
          <td><%=request("publishtime")%></td>
        </tr>
  </table>
  <p><a href="first.asp">�����ϴ�ҳ</a></p>

</div>
<!-- InstanceEndEditable -->
<p align="center"><font color="#CCFF66">���Ǿ��ֲ�����Ȩ���� <font face="Arial, Helvetica, sans-serif"><br>
  Copyright 2002-2003 Clustering Intelligence Club</font></font> </p>
</body>
<!-- InstanceEnd --></html>
