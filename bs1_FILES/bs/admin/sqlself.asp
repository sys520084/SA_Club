<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<%if not session("admin") then
   Response.Write "<div align=center>对不起，您不能打开本页！</div>"
   Response.End
  end if
  sql=request("sql")
  set rs=server.CreateObject("ADODB.Recordset")
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  conn.BeginTrans
  conn.Errors.Clear 
  on error resume next
  if left(sql,6)="select" then
     res=true
     rs.Open sql,conn,3
  else
    conn.Execute sql
  end if
  for i=0 to conn.Errors.Count-1
     Response.Write conn.Errors(i).Description&"<br>"
  next
  if res then
     do while not rs.EOF%><div><%
        for i=0 to rs.Fields.Count-1%>
           <%=rs.fields(i).name%>=<%=rs.Fields(i).Value%>&nbsp;
        <%next
     rs.MoveNext%>
     </div>
    <%loop
    rs.Close
  end if
  if conn.Errors.Count>0 then
     conn.RollbackTrans
  else
     conn.CommitTrans
  end if      
%>
<P>语句执行成功</P>        
</BODY>
</HTML>
