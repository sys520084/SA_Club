<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<%if not session("admin") then
     Response.Write "<div algin='center'>对不起，您不能打开本页。<a href='jingwen.asp'>请重新登陆</a>"
     response.end
  end if
  id=request("id")
  if id="" then
     Response.Write "没有指定要删除的纪录！"
     response.end
  end if
    'DbPath=server.MapPath("/music")&"/club/db/apply.mdb"
    set conn=server.CreateObject("ADODB.Connection")
    conn.Open OpenString
    sql="delete from application where id="&id
    conn.Execute sql
  %>     
<P>成功删除这条纪录。</P>

</BODY>
</HTML>
