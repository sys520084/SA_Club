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
<title>ɾ����Ϣ</title>
<BODY bgColor=white>


<P>&nbsp;</P>
<%if session("user")="" then
     Response.Write "<div align='center'>�Բ�������Ҫ��ͨ�������֤��</div>" 
     Response.End
  end if
  if request("id")="" then
     Response.Write "<div  align='center'>�Բ�����û��ѡ��Ҫɾ���ļ�¼��</div>"
     Response.End
  end if
  id=request("id")
  set conn=server.CreateObject("ADODB.Connection")
  Conn.open OpenString
  sql="delete from message where id="&id
  conn.Errors.Clear
  conn.Execute(sql)
  if conn.Errors.Count>0 then
     Response.Write "�Բ���ɾ����Ϣʱ�������ܵ�ԭ���ǣ�������Ϣ�Ѿ��������ˡ�"
     conn.Errors.Clear 
     Response.End
  end if
%>
<p align="center">ɾ����Ϣ�ɹ�����������ܿ���������Ϣ���밴������ϵ�ˢ�°�ť��</p>
<p align="center"><a href="javascript:window.close();">�رմ˴���</a></p>     
    
</BODY>
</HTML>
