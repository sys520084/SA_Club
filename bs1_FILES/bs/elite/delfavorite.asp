<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<%username=trim(session("user"))
if username="" or isnull(username) then
	response.write "�㻹û�е�½�����ܴ򿪱�ҳ��"
	response.end
end if
%>
<title><%=username%>���ղ�Ŀ¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
DIV,p {FONT-FAMILY: "����"; FONT-SIZE: 9pt; color:#000000;}
TD {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
}
A {COLOR: #0099FF; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #FF0000; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
input,textarea,SELECT {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
	background-color: #FFFFFF;
}</STYLE>
</head>

<body>
<p>
  <%
set conn=server.CreateObject("ADODB.Connection")
Conn.open OpenString
tb=1
if request("id")<>"" then
	sql="delete from favorite where article_id="&request("id")
	tb=2
	conn.execute sql
else
	response.write "�Բ�����û��ѡ��ȥ�ղص�����"
	response.end 
	
end if

%>
<div align=center>�Ѿ��ɹ�ɾ����¼</div>

</body>
</html>
