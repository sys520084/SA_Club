<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<style>
<!--
div,td,p{font-size:9pt; line-height:14pt; font-family:����;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
-->
</style>
</HEAD>
<title>������Ϣ</title>
<BODY bgColor=white>


<P>���е���Ϣ���£���������������Ϊ���ܱ��汾ҳ��</P>
<%if session("user")="" then
     Response.Write "<div align='center'>�Բ�������Ҫ��ͨ�������֤��</div>" 
     Response.End
  end if
  set conn=server.CreateObject("ADODB.Connection")
  Conn.open OpenString
  sql="select * from message where reciever='"&session("user")&"' order by id desc"
  conn.Errors.Clear
  set rs=conn.Execute(sql)
  do while not rs.eof%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr> 
    <td width="50%">�����ˣ�<%=rs("sender")%></td>
    <td width="50%">�������ڡ�ʱ�䣺<%=rs("senddate")&rs("sendtime")%></td>
  </tr>
  <tr> 
    <td colspan="2">���⣺<%=rs("title")%></td>
  </tr>
  <tr bgcolor="#E1E1F7"> 
    <td colspan="2"><%=rs("content")%></td>
  </tr>
</table>
</br>
<%	rs.movenext
   loop
if conn.Errors.count>0 then
     Response.Write "�Բ��𣬱�����Ϣʱ�����"
     conn.Errors.Clear 
     Response.End
  end if
%>   
    
</BODY>
</HTML>
