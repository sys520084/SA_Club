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
<title><%=username%>�ľ���Ŀ¼</title>
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
if request("categoryid")<>"" and request("articleid")<>"" then
	sql="delete from article_category where article_id="&request("articleid")&" and category_id="&request("categoryid")
	tb=2
else
	sql="delete from article_category where id="&request("relationid")
end if
conn.execute sql
%>
<div align=center>�Ѿ��ɹ�ɾ����¼</div>
<%if tb=1 then%>
<script language=javascript>
window.location.replace("showdir.asp?articleid=<%=request("articleid")%>");
</script>
<%end if%>
</body>
</html>
