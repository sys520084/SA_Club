<%@ Language=VBScript %>
<!--#include file="../../var.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<%catagory=request("catagory")
if (not session("pass"&catagory)) and not session("admin") then
   Response.Write "<div align=center>�Բ��������ܴ򿪱�ҳ��</div>"
   Response.End
end if
if isnull(Request("id")) or Request("id")="" then
   Response.Write "<div align=center>�Բ�����������˴򿪱�ҳ��</div>"
   Response.End
end if
id=request("id")
username=request("username")
set conn=server.CreateObject("ADODB.Connection")
conn.Open OpenString
sql="delete from power where id="&id
conn.Execute sql
%>
<P align=center>�ɹ��Ľ���˱���յģ�<%=username%>��Ȩ�ޡ�</P>
</BODY>
</HTML>
