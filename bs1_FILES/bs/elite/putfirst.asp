<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<!--#include file="../forum/aspFunctions.asp"-->
<%
OpenString=DBBS()%>
<html>
<head>
<%username=trim(session("user"))
if username="" or isnull(username) then
	response.write "你还没有登陆，不能打开本页！"
	response.end
end if
set conn=server.CreateObject("ADODB.Connection")
conn.Open OpenString
isadmin=IsSuperAdmin(Conn,Session("user"))

if isadmin then
	id=request("id")
	sql="select * from articles where id="&id
	set rs=server.CreateObject("ADODB.Recordset")
	'response.write sql
	'response.end
	rs.open sql,conn,1,3
	 
	if rs.eof then
		response.write "该文章不属于精华，不能放置在首页"
		response.End()
	end if
	score=rs("score")
	
	if isnull(score) then
	response.write score&"ok"
		rs("score")=1
	else
	 rs("score")=1-rs("score")
	end if
	 rs.update
	 rs.close
	 set rs=nothing
end if
	 set conn=nothing
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
 <div align=center>已经成功添加到了首页</div>

</script>
</body>
</html>
