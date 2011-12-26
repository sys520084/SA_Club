<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>选择用户名</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:black;}
a:link {
	text-decoration: none;
	color: #003366;
}
a:visited {
	text-decoration: none;
	font-family: "宋体";
	font-size: 9pt;
	color: #000099;
}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;
	color: #000066;
}
-->
</style>
<script>
function givevalue(val)
{
	window.returnValue=val;
	window.close();
}
</script>
</head>
<body bgcolor="#FFFFFF">
<%  
if session("user")="" then
     Response.Write "<div align='center'>对不起，您需要先通过身份验证。</div>" 
     Response.End
  end if
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  'Response.End 
   for i=65 to 91
		if i<=90 then
			toPrint=toPrint&"<a href='#"&chr(i)&"'>"&chr(i)&"</a>&nbsp;"
		else
			toPrint=toPrint&"<a href='#其它'>其它</font></a>&nbsp;"
		end if
   next
   Response.Write toPrint	
   cols=5
   for i=65 to 91
	if i<=90 then
		sql="select * from users where username like '"&chr(i)&"%' order by username"
	else
		sql="select * from users where asc(username)<65 order by username"
	end if
   ' Response.Write sql
    'Response.End 
	set rs=conn.Execute(sql)
	j=0%>
	<table width=100% border=0 align="center" cellpadding="0" cellspacing="0" bgcolor="#6699FF">
	<tr>
	<%if i<=90 then%>
    <td width=12% bgcolor="#FFFFFF"><a name="<%=chr(i)%>"></a><%=chr(i)%></td><%for j=1 to cols-1%><td width=12% bgcolor="#FFFFFF">&nbsp;</td><%next%>
    <%else%>
    <td width=12% bgcolor="#FFFFFF"><a name="其它"></a>其它</td><%for j=1 to cols-1%><td width=12% bgcolor="#FFFFFF">&nbsp;</td><%next%>
    <%end if%>
	</tr>
	<%do while not rs.eof
	 for j=1 to cols
		if not rs.eof then
			name=rs("username")
		end if
	    if j=1 then%> 
			<tr><%end if%>
	    <td width=12% align=center><a href=# onclick="javascript: givevalue('<%=name%>')"><%=name%></a></td>
	    <%if j=cols then%> 
			</tr><%end if%>
	<%
		if not rs.eof then
			rs.movenext
		end if
	 next
	 loop
	 rs.close%>
	 </table><table width=100%><tr><td width=100% bgcolor=white><a href="#top">△Top</a></td></tr></table>
<%   next%>  

</body>
</html>
