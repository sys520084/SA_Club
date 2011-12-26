<%@ Language=VBScript %>
<!--#include file="../VAR.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
 <%
  set conn=server.CreateObject("ADODB.Connection")
  set rs=server.CreateObject("ADODB.Recordset")
  sql="select * from users"
  conn.Open OpenString
  rs.Open sql,conn,1, 3
  conn.BeginTrans 
  do while not rs.EOF
     username=rs("username")
     sql="SELECT COUNT(*) AS score FROM articles WHERE (score = 1) AND (author LIKE '"&username&"(%' or author='"&username&"')"
     set rs1=conn.Execute(sql)
     if not rs1.eof then
        rs("markscore")=cint(rs1("score"))
     end if
     rs.MoveNext
  loop
  for i=0 to conn.Errors.Count-1
      Response.Write conn.Errors(i).Description
  next
  if not conn.Errors.Count=0 then
     Response.Write "有错误!"
     conn.RollbackTrans
     Response.End
  end if   
  'conn.RollbackTrans
  conn.CommitTrans 
  %>    
     
<P>更新完成111</P>
</BODY>
</HTML>
