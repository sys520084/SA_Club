<%@ Language=VBScript %>
<!--#include file="..\..\var.asp"-->
<%OpenString=DBBS()%>
<%
  user=session("user")
  id=request("id")
  sql="select * from files where id="&id
  set conn=server.CreateObject("ADODB.Connection")
  set rs=server.CreateObject("ADODB.Recordset")
  'Response.Write sql
  'Response.End 
  conn.Open OpenString
  rs.Open sql,conn,1,3
  on error resume next
  if not rs.EOF then
	filename=rs("filename")
    if session("user")<>rs("username") then
		rs("times")=rs("times")+1
		rs.Update
	end if
  end if
  rs.Close 
  on error goto 0
  'Response.Write filename
  'Response.End 
  if left(filename,len("http://"))<>"http://" then
  Response.Redirect "../files/"&filename
  else
  response.Redirect filename
  end if
  %>
  