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
  'Response.Write openstring
  'Response.End 
  conn.Open OpenString
  set rs=server.CreateObject("ADODB.Recordset")
  sql="select * from article_category order by id desc"
  rs.open sql,conn,3
  do while not rs.eof
	articleid=rs("article_id")
	sql1="select * from articles where id="&articleid
	set rs1=server.CreateObject("ADODB.Recordset")
	rs1.open sql1,conn,1,3
	rs1("score")=1
	rs1.update
	response.write rs1("title")&"<br>"
	rs1.close
	set rs1=nothing
	
	'sql="update articles set author='"&author&"' where id="&rs("id")
	'conn.execute sql
	rs.movenext
  loop
  rs.close
  set rs=nothing
  %>
<P>&nbsp;</P>

</BODY>
</HTML>
