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
set rs=server.CreateObject("ADODB.Recordset")
set conn=server.CreateObject("ADODB.Connection")
Conn.open OpenString
%>
<%categoryid=request("categoryid")
fatherid=request("fatherid")
if fatherid="" or isnull(fatherid) then
	fatherid=-1
end if
if request("categoryname")<>"" then
	sql="select * from usercategory where username='"&username&"'"
	rs.open sql,conn,1,3
	rs.addnew
	rs("username")=username
	rs("category")=request("categoryname")
	rs("descript")=request("descript")
	rs("fatherid")=fatherid
	rs.update
	categoryid=rs("id")
	rs.close
%>
  <p align="center">Ŀ¼�����ɹ�</p>
<%end if
if request("articleid")<>"" then
	set rstemp=server.CreateObject("ADODB.Recordset")
	if request("ocategoryid")<>"" then
		sql="delete from article_category where article_id="&request("articleid")&" and category_id="&request("ocategoryid")
		conn.execute sql
	end if
	sql="select * from articles where id="&request("articleid")
	rstemp.open sql,conn,3
	if not rstemp.eof then
		'response.write rstemp("author")
		'response.write username
		if rstemp("author")<>username then
			response.write "<div align=center>�㲻����������˵����µ���ľ���Ŀ¼��</div>"
			response.end
		end if
	end if
	rstemp.close
	sql="select * from article_category where article_id="&request("articleid")&" and category_id="&categoryid
	'response.write sql
	'response.end
	rstemp.open sql,conn,3
	if not rstemp.eof then
		response.write "<div align=center>���ܰ�ͬһƪ���·ŵ�ͬһ��Ŀ¼��</div>"
		response.end
	end if
	rstemp.close
	set rstemp=nothing
	sql="insert into article_category (article_id,category_id) values ("&request("articleid")&","&categoryid&")"
	'response.write sql
	'response.end
	conn.execute sql%>
  <p align="center">������ӳɹ�</p>
<%end if%>
</form>
<script language="javascript">
window.close();
</script>
</body>
</html>
