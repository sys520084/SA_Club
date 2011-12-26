<!--#include file="../../../VAR.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:宋体;color:black;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0000CC}
a:visited {text-decoration: none;; color: #660066}
a:hover {text-decoration: underline; color: #FF0000}
a:active {text-decoration;background-color: #CCCCCC; color: #006666}
-->
</style>
<title>修改成功</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<%catagory=request("catagory")
  id=request("id")
  welcome=request("welcome")
  intro=request("intro")
  if id="" then
    Response.Write "<div align=center>对不起，不能进行更改！</div>"
    Response.End
  end if  
  id=cint(id)
  if (not session("pass"&catagory)) and not session("admin") then
   Response.Write "<div align=center>对不起，您不能打开本页！</div>"
   Response.End
  end if
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  set rs=server.CreateObject("ADODB.Recordset")
  sql="select * from forum_state where id="&id
  rs.Open sql,conn,1,3
  if rs.EOF then
     Response.Write "<div align=center>没有找到指定的论坛</div>"
     Response.End
  end if   
  %>
<body bgcolor="#FFFFFF">
<div align="center">
  <%if len(welcome)>500 then%>
  <p>出错，太长的文字：欢迎文字。</p>
  <%Response.End 
    end if
   if len(intro)>200 then%>
     <p>出错，太长的文字：介绍文字。</p>
  <%Response.End 
   end if
 rs("welcome")=welcome
 rs("intro")=intro
 rs.Update
 rs.Close
   %>
  <p>论坛状态修改成功</p>
  <p>&nbsp;</p>
</div>
</body>
</html>
