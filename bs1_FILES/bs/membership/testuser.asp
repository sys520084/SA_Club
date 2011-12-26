<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<%OpenString=DBBS()
function check(str)
	if (asc(str)>=48 and asc(str)<=57) or (asc(str)>=65 and asc(str)<=122) or asc(str)<0 then
		check=true
	else
		check=false
	end if
end function
set Conn=server.CreateObject("ADODB.Connection")
Conn.open OpenString
set rs=server.CreateObject("ADODB.Recordset")
username=request("username")
%>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="MSHTML 5.00.2920.0" name=GENERATOR>
<META content=none name="Microsoft Border">
<STYLE type=text/css>
TD {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
}
div {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
}
A {COLOR: #0066FF; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #FF0000; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
input,SELECT {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
	background-color: #FFFFFF;
}
</STYLE>
</HEAD>
<BODY bgColor=#FFFFFF text="#FFFFFF">
<%
if username="" then
	response.write "<div align=center>对不起，用户名不能为空！请重试</div>"
	response.end
end if
for i=1 to len(username)
	ch=mid(username,i,1)
	if not check(ch) then
		response.write "<div align=center>对不起，用户名中只能含有字符或者数字！请重试</div>"
		response.end
		exit for
	end if
next

sql="select * from users where username='"&username&"'"
rs.open sql,conn,3
if not rs.eof then
%>
 <DIV align=center>对不起，<font color="red"><%=username%></font>这个名字已经有人申请了！</DIV>
<%else%>
 <DIV align=center><font color="green">恭喜，您可以使用<font color="red"><%=username%></font>这个用户名！</font></DIV>
<%end if
rs.close
set rs=nothing
%>
</BODY></HTML>
