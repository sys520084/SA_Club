<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<%if not session("admin") then
     Response.Write "<div algin='center'>�Բ��������ܴ򿪱�ҳ��<a href='jingwen.asp'>�����µ�½</a>"
     response.end
  end if
  id=request("id")
  if id="" then
     Response.Write "û��ָ��Ҫɾ���ļ�¼��"
     response.end
  end if
    'DbPath=server.MapPath("/music")&"/club/db/apply.mdb"
    set conn=server.CreateObject("ADODB.Connection")
    conn.Open OpenString
    sql="delete from application where id="&id
    conn.Execute sql
  %>     
<P>�ɹ�ɾ��������¼��</P>

</BODY>
</HTML>
