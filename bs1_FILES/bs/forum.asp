<!--#include file="../VAR.asp"-->
<html>
<head>
<title>ͷ�Է籩����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
-->
</style>
</head>
<%
user=request("user")
  if trim(session("user"))="" then
     'Response.Write "<div align='center'>�Բ��������ܴ򿪱�ҳ��</div>"
     'Response.end
	 user=""
  end if   %>
     
<frameset cols="160,*" frameborder="NO" border="0" framespacing="0"> 
  <frame name="leftFrame" scrolling="ifneeded" src="left.asp">
  <frame name="mainFrame" src="membership/bsblog.asp?open=1&user=<%=session("user")%>">
</frameset>
<noframes><body bgcolor="#FFFFFF">

</body></noframes>
</html>
