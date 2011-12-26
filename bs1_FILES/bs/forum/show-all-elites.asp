<!--#include file="../../var.asp"-->
<!--#include file="convert.inc"-->
<!--#include file="aspFunctions.asp"-->
<!--#include file="../publicmodule/showtitles.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>察看所有精华</title>
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:"宋体";color:#333333;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0066FF}
a:visited {text-decoration: none;; color: #0033CC}
a:hover {text-decoration: underline; color: #0033CC}
a:active {text-decoration;color: #0066FF}
.art {
	font-size: 10pt;
	color: #0000CC;
	font-style: normal;
}
-->
</style>
<body bgcolor="#FFFFFF" text="#CCCCCC" link="#0000CC" vlink="#000000">
<%set conn=server.CreateObject("ADODB.Connection")
  Conn.open OpenString

  dim result
  dim resV
  dim sql

  '返回搜索结果
  'response.write request("order")
  result=SearchingTopics("articles")
  resV=split(result,vbTab)
  sql=resV(0)
  

  UrlTail=""
  'response.write request("subs")
  isadmin=IsSuperAdmin(Conn,Session("user"))
  set rs=server.CreateObject("ADODB.Recordset")

  page=request("page")
  'response.write page
  if page="" then
	page=1
  else
	page=cint(page)
  end if
  'Response.Write sql
   'Response.End 
  rs.Open sql,conn,3
  'Response.End     
  pgSz=20
  rs.PageSize=pgSz
  'response.write page
  'Response.Write request("flag")
  if not rs.EOF then
      rs.AbsolutePage=page
  end if
  pgCount=rs.PageCount
  rsCount=rs.RecordCount
  if isnull(rsCount) then
     rsCount=0
  end if   
  if isnull(pgCount) then
     pgCount=0
  end if

showPages "show-all-elites.asp",UrlTail,"",1
showTitles "viewblog.asp","new.asp","../membership/display.asp",UrlTail,"",rs,"show-all-elites.asp"
showPages "show-all-elites.asp",UrlTail,"",2%>
  <div>
  <p align=right>
<%if isadmin then%>
      <a href="../admin/admin.asp"><font color="blue">管理</font></a>
<%end if%> </p>
    <div align="center"> 
      <p><font color="#333333">集智俱乐部·版权所有 <font face="Arial, Helvetica, sans-serif"><br>
        Copyright 2007-2027 Swarm Agents Club</font></font></p>
    </div>
  </div>

</body>
</html>