<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<%if not session("admin") then
     Response.Write "<div align=center>�Բ��������ܷ��ʱ�ҳ��</div>"
     Response.End
  end if 
  username=request("username")
  if username<>"" then
     set conn=server.CreateObject("ADODB.Connection")
     conn.Open OpenString
     conn.BeginTrans
     on error resume next
     sql="delete from power where username='"&username&"'"
     conn.Execute(sql)
     sql="delete from detail  where username='"&username&"'"
     conn.Execute(sql)
     sql="delete from memberdes where username='"&username&"'"
     conn.Execute(sql)
     sql="delete from users where username='"&username&"'"
     conn.Execute(sql)
     conn.CommitTrans
  %>
<P>�ɹ�ɾ���û���<%=username%></P>
<%else%>
<p>û��ָ��Ҫɾ�����û�����</p>
<%end if   %>
</BODY>
</HTML>
