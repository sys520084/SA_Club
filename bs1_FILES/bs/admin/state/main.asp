<!--#include file="../../../VAR.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>高级论坛管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:宋体;color:black;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0000CC}
a:visited {text-decoration: none;; color: #660066}
a:hover {text-decoration: underline; color: #FF0000}
a:active {text-decoration;background-color: #CCCCCC; color: #006666}
-->
</style>
</head>

<body bgcolor="#FFFFFF">
<%catagory=request("catagory")
  if (not session("pass"&catagory)) and not session("admin") then
   Response.Write "<div align=center>对不起，您不能打开本页！</div>"
   Response.End
  end if
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  set rs=server.CreateObject("ADODB.Recordset")
  sql="select * from forum_state where catagory='"&catagory&"'"
  rs.Open sql,conn,3
  if rs.EOF then
     Response.Write "<div align=center>对不起，没找到这个论坛。</div>"
     Response.End
  end if
  %>   
<div align="center">
  <p>高级论坛管理:<font color=red><%=catagory%></font></p>
  <form name="form1" method="POST" action="modifystate.asp">
    <p>定制欢迎文字： </p>
    <p> 
      <textarea name="welcome" cols="50" rows="5"><%=trim(rs("welcome"))%></textarea>
      <br>
    </p>
    <p>定制介绍文字：</p>
    <p> 
      <textarea name="intro" cols="50" rows="5"><%=trim(rs("intro"))%></textarea>
    </p>
    <p> 
      <input type=hidden name="id" value=<%=rs("id")%>>
      <input type=hidden name="catagory" value="<%=catagory%>">
      <input type="submit" name="confirm" value="确定">
      <input type="reset" name="reset" value="重写">
    </p>
    </form>
</div>
</body>
</html>
