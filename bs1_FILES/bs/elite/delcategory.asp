<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<%username=trim(session("user"))
if username="" or isnull(username) then
	response.write "你还没有登陆，不能打开本页！"
	response.end
end if
%>
<title><%=username%>的精华目录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
DIV,p {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; color:#000000;}
TD {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
}
A {COLOR: #0099FF; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #FF0000; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
input,textarea,SELECT {
	FONT-FAMILY: "宋体";
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
categoryid=request("categoryid")
sql="select * from usercategory where username='"&username&"' and id="&categoryid
rs.open sql,conn,1,3
if rs.eof then
	response.write "<div align=center>对不起，你不能修改其它人的目录</div>"
	response.end
end if
rs.close
set rs=nothing
sql="delete from usercategory where username='"&username&"' and id="&categoryid
conn.execute sql
sql="delete from article_category where category_id="&categoryid
conn.execute sql
sql="update usercategory set fatherid=-1 where username='"&username&"' and fatherid="&categoryid
conn.execute sql
%>
  <p align="center">目录删除成功</p>
</body>
</html>
