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
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
}
div {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
}
A {COLOR: #0066FF; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #FF0000; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
input,SELECT {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
	background-color: #FFFFFF;
}
</STYLE>
</HEAD>
<BODY bgColor=#FFFFFF text="#FFFFFF">
<%
if username="" then
	response.write "<div align=center>�Բ����û�������Ϊ�գ�������</div>"
	response.end
end if
for i=1 to len(username)
	ch=mid(username,i,1)
	if not check(ch) then
		response.write "<div align=center>�Բ����û�����ֻ�ܺ����ַ��������֣�������</div>"
		response.end
		exit for
	end if
next

sql="select * from users where username='"&username&"'"
rs.open sql,conn,3
if not rs.eof then
%>
 <DIV align=center>�Բ���<font color="red"><%=username%></font>��������Ѿ����������ˣ�</DIV>
<%else%>
 <DIV align=center><font color="green">��ϲ��������ʹ��<font color="red"><%=username%></font>����û�����</font></DIV>
<%end if
rs.close
set rs=nothing
%>
</BODY></HTML>
