<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><!-- InstanceBegin template="/Templates/bs.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" --> 
<title>Untitled Document</title>
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
.unnamed1 {  font-family: "宋体"; font-size: 9pt; font-style: normal; line-height: 100%; font-weight: normal; color: #99FFFF}
.unnamed2 {  font-family: "宋体"; font-size: 9pt; line-height: 100%; font-weight: normal; color: #294C39}
input {
	font-family: "宋体";
	font-size: 10pt;
	color: #FFFFFF;
	background-color: #000033;
}
textarea {
	font-family: "宋体";
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
  <p>在线用户列表：</p>
  <%if session("user")="" then
		response.write "<div align=center>对不起，您不能使用本功能，请先登陆！</div>"
		response.end
	end if
	time1=dateadd("s",-300,now())
	sql="select * from online where onlinetime>#" &time1&"# order by onlinetime desc,username"
	set conn=server.CreateObject("ADODB.Connection")
	conn.open openstring
	set rs=server.CreateObject("ADODB.Recordset")
	'response.write sql
	'response.end
	'rs.open sql,conn,3
	%>
	
  <table width="75%" border="1">
    <tr bgcolor="#3366FF"> 
      <td width="16%"><div align="center"><font color="#000000">用户</font></div></td>
      <td width="16%"><div align="center"><font color="#000000">操作</font></div></td>
      <td width="1" bgcolor="#000000"></td>
      <td width="16%"><div align="center"><font color="#000000">用户</font></div></td>
      <td width="16%"><div align="center"><font color="#000000">操作</font></div></td>
      <td width="1" bgcolor="#000000"></td>
      <td width="16%"><div align="center"><font color="#000000">用户</font></div></td>
      <td width="16%"><div align="center"><font color="#000000">操作</font></div></td>
    </tr>
    <%ti=0
    do while not rs.eof%>
    <tr>
	  <% for i=1 to 3
			if not rs.eof then
				user=rs("username")
				write="<A href='display.asp?name="&user&"' onClick='return js_callpages(this.href)'>"&user&"</a>"
			else
				user=""
				write="&nbsp;"
			end if%>
      <td><div align="center"><%=write%></div></td>
      <td><div align="center"><%if user<>"" then%><a href="../message/sendmsg.asp?name1=<%=trim(user)%>"><img src="../img/envolop.gif" alt="给他（她）发消息" width="15" height="13" border="0"></a><%else%>&nbsp;<%end if%></div></td>
      <%if i<3 then%>      <td width="1" bgcolor="#000000"></td><%end if%>

      <%	if not user="" then
				rs.movenext
			end if
		next%>
    </tr>
	<%loop
	rs.close%>
  </table>
  <p>&nbsp;</p>
</div>
<!-- InstanceEndEditable -->
<p align="center"><font color="#CCFF66">集智俱乐部·版权所有 <font face="Arial, Helvetica, sans-serif"><br>
  Copyright 2002-2003 Clustering Intelligence Club</font></font> </p>
</body>
<!-- InstanceEnd --></html>
