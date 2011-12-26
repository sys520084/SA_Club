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
set rs=server.CreateObject("ADODB.Recordset")
set conn=server.CreateObject("ADODB.Connection")
Conn.open OpenString

function ShowSubTree(byval id)
	sql="select * from usercategory where fatherid="&id
	   set rs1=conn.Execute(sql)
	do while not rs1.eof
		Tid=rs1("id")
		Tcatalog_name=trim(rs1("category"))
		Tfather_id=trim(rs1("fatherid"))%>
</p>
	<ul>
    <li><a href=# onclick="javascript: givevalue('<%=Tid%>')"><%=Tcatalog_name%></a> 
      <%ShowSubTree Tid%>
    </li>
  </ul>
  <%rs1.movenext
	loop
	rs1.close
	set rs1=nothing
end function
sql="select * from articles where id="&request("articleid")
rs.open sql,conn,3
if not rs.eof then
	title=rs("title")
end if
rs.close

if request("categoryid")="" then
sql="select * from article_category,usercategory where article_id="&request("articleid") &" and usercategory.id=article_category.category_id"
'response.write sql
'response.end
rs.open sql,conn,3
if not rs.eof then
%>
<div align=left>文章<font color=red><%=title%></font>已经属于目录：</div>
<ul>
<% do while not rs.eof%>
<li><div><%=rs("category")%><a href="delElite.asp?relationid=<%=rs(0)%>&articleid=<%=request("articleid")%>"><img border=0 src="../images/delete.gif" alt="从这个目录中删除这篇文章"></a></div></li>
<% rs.movenext
loop%>
</ul>
<hr width=90% align=center>
<%end if
rs.close
end if%>
<div align=left>将文章<font color=red><%=title%></font>归入目录：（点击下面的目录名称）</div>

<%

sql="Select * From usercategory where fatherid=-1 and username='"&username&"'"
  'Response.Write sql 
  rs.Open sql,Conn,3
  if rs.eof then%>
  <p align="center">您当前没有任何精华目录</p>
<%end if
   do while not rs.eof 
  		Tid=trim(rs("id"))
		Tcatalog_name=trim(rs("category"))
		Tfather_id=trim(rs("fatherid"))%>
  <ul>
    <li><a href=# onclick="javascript: givevalue('<%=Tid%>')"><%=Tcatalog_name%></a> 
      <%ShowSubTree Tid%>
    </li>
  </ul>
  <p> 
    <% rs.movenext
	  loop
  %>
  </p>
  <form action="addware.asp" method="post" name="mainfrm" id="mainfrm">
  <input name="categoryid" type="hidden" value="">
  <input name="ocategoryid" type="hidden" value="<%=request("categoryid")%>">
  <input name="articleid" type="hidden" value="<%=request("articleid")%>">
</form>
<hr width=90% align=center>
<div align=left>新建精华目录，并把<font color=red><%=title%></font>归入新建的目录</div>
<form name="newForm" method="post" action="addware.asp">
<%
rs.close
sql="select * from usercategory where username='"&username&"'"
rs.open sql,conn,3
%>
  <p align="center">在 
    <select name="fatherid" id="fatherid">
      <option value="-1">根目录</option>
      <%do while not rs.eof%>
      <option value="<%=rs("id")%>"><%=rs("category")%></option>
      <%rs.movenext
loop
rs.close
set rs=nothing%>
    </select>
    目录下，新建一个目录: 
    <input type="text" name="categoryname">
    <input type="hidden" name="articleid" value="<%=request("articleid")%>">
   <input name="ocategoryid" type="hidden" value="<%=request("categoryid")%>">
	<input type="hidden" name=
  </p>
  <p align="center">新建目录介绍：</p>
  <p align="center">
    <textarea name="descript" cols="50" rows="5" id="descript"></textarea>
  </p>
  <p align="center">
    <input type="submit" name="Submit" value="Submit">
  </p>
  </form>
<p align="center">&nbsp;</p>
</body>
</html>
