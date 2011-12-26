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
%>
<%
categoryid=request("categoryid")
if request("confirm")="确定" then
	fatherid=request("fatherid")
	descript=request("descript")
	if fatherid="" or isnull(fatherid) then
		fatherid=-1
	end if
	if request("categoryname")<>"" then
		sql="select * from usercategory where username='"&username&"' and id="&categoryid
		rs.open sql,conn,1,3
		if rs.eof then
			response.write "<div align=center>对不起，你不能修改其它人的目录</div>"
			response.end
		end if
		rs("category")=request("categoryname")
		rs("fatherid")=fatherid
		rs("descript")=descript
		rs.update
		categoryid=rs("id")
		rs.close
	else
		response.write "<div align=center>目录名不能为空，请<a href='javascript:history.go(-1);'>返回</a>"
		response.end
	end if
%>
  <p align="center">目录修改成功</p>
<%else

sql="select * from usercategory where username='"&username&"' and id="&categoryid
		rs.open sql,conn,1,3
		if rs.eof then
			response.write "<div align=center>对不起，你不能修改其它人的目录</div>"
			response.end
		end if
sql1="select * from usercategory where username='"&username&"'"
set rstemp0=server.CreateObject("ADODB.Recordset")
rstemp0.open sql1,conn,3
%>
<form name="frm" action="editcategory.asp" method="post">
	<input type="hidden" name="categoryid" value="<%=categoryid%>">
  <p align="center">修改该目录的父目录为：
    <select name="fatherid" id="fatherid">
	<option value="-1">根目录</option>
<%do while not rstemp0.eof
	if rstemp0("id")<>clng(categoryid) then%>
	<option value="<%=rstemp0("id")%>" <%if rstemp0("id")=rs("fatherid") then%>selected<%end if%>><%=rstemp0("category")%></option>
<%  end if
rstemp0.movenext
loop
rstemp0.close
set rstemp0=nothing%>
    </select>
<p align=center>修改目录名: 
    <input type="text" name="categoryname" value="<%=rs("category")%>"></p>
	  <p align="center">目录介绍：</p>

<p align="center">
    <textarea name="descript" cols="50" rows="5" id="descript"><%=rs("descript")%></textarea>
  </p>
<p align="center">
    <input type="submit" name="confirm" value="确定"></P>
  </p>
  </form>
<p align="center">&nbsp;</p>
<%end if%>
</form>
</body>
</html>
