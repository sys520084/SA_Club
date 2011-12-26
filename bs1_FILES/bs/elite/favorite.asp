<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<%username=session("user")
if username="" or isnull(username) then
	response.write "你还没有登陆，不能打开本页！"
	response.end
end if
%>
<title><%=username%>的收藏</title>
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
}.style1 {color: #FF0000}
.style2 {color: #000000}
</STYLE>
<script>
function givevalue(val)
{
	mainfrm.categoryid.value=val;
	mainfrm.submit();
}
</script>
</head>

<body>
<p>
  <%
id=request("id")
if id="" then
	response.write "您没有指定要收藏的文章，不能操作！"
	response.End
end if  
if session("user")="" then
		 UrlTail=MyUrlEncode("../elite/favorite.asp?id=" & request("id"))
         Response.Write "<div align='center'>对不起，您不能发表文章，请<a href='../membership/relogin.asp?UrlTail=" & UrlTail & "'>重新登陆</a>，或者：</div>"
         Response.End
 end if
set rs=server.CreateObject("ADODB.Recordset")
set conn=server.CreateObject("ADODB.Connection")
Conn.open OpenString
if request("submit")<>"确定" then
	sql="select * from favorite where article_id="&Id
	'response.write sql
	'response.end 
	rs.open sql,conn,3
	if not rs.eof then
		response.write "您已经收藏这篇文章了，不能重复收藏！"
		response.End
	end if
	rs.close
	sql="select * from articles where id="&id
	'response.write sql
	'response.end 
	rs.open sql,conn,3
	'response.end
	if rs.eof then
		response.write "没有找到您要收藏的文章，请确认此文章是否已经被删除？"
		response.End
	end if
	
	title=rs("title")
	author=rs("author")
	rs.close
	set rs=nothing
%>
<form name="form1" method="post" action="">
  <p>你要收藏<span class="style1"><%=author%><span class="style2">的</span><%=title%></span></p>
  <p>请为你要收藏的这篇文章添加标签说明，标签以空格隔开:</p>
  <p>
    <input name="id" type="hidden" id="id">
  </p>
  <p>
    <input name="tags" type="text" id="tags" size="50">
    <input type="submit" name="Submit" value="确定">
    </p>
</form>
<%
else
	tags=MyhtmlEncode(request("tags"))
	username=session("user")
	ondate=date()&" "&time()
	sql="insert into favorite (username,article_id,tag,ondate) values ('"&username&"',"&id&",'"&tags&"',#"&ondate&"#)"
	'response.write sql
	conn.execute(sql)
	conn.close
	set conn=nothing
	%>
您已经成功收藏1篇文章！
	
	<%
end if
%>
<p><p align="center">&nbsp;</p>
</body>
</html>
