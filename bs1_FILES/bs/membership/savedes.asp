<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>���ϸ���</title>
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline;}
a:active {text-decoration;}
-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></HEAD>
<BODY bgcolor="#FFFFFF" link="#0033CC" vlink="#0033CC" alink="#0033CC">
<%if session("user")="" then
     Response.Write "<div align=center>�Բ�����û�з��ʱ�ҳ��Ȩ�ޣ�</div>"
     Response.End
  end if
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  conn.BeginTrans 
  set rs=server.CreateObject("ADODB.Recordset")
  name=session("user")
  on error resume next
   sql="select * from memberdes where username='"&name&"'"
   'Response.Write sql&request("username")
   'response.end
   rs.Open sql,conn,1,3
   if rs.EOF then
     Response.Write"<div align=center>�Բ���û���ҵ����û���</div>"
     Response.end
   end if
   rs("showtitle")=request("showtitle")
   rs("readme")=request("story")
   rs("detailcontent")=request("detailcontent")
   rs.Update
   if conn.Errors.Clear>0 then
      Response.Write "<div align=center>�������󣬿�������ΪĳЩ���ݳ�����</div>"
      for i=0 to conn.Errors.Count-1
         Response.Write conn.Errors(i).Description&"<br>"
      next   
      Response.End
   end if
   conn.CommitTrans
%>
<P align=center>�ɹ��������ĸ���˵�����ϡ�</P>
<P align=center><a href="../forum.asp?user=<%=session("user")%>"" target="_parent">���������</a></P>

</BODY>
</HTML>
