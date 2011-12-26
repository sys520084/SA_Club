<!--#include file="../../../VAR.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>选择文件夹</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
-->
</style>
<script>
function givevalue(val)
{
	window.returnValue=val;
	window.close();
}
</script>
</head>
<body bgcolor="#FFFFFF">
<%Sub ShowSubTree(byval directory,byval catagory)
   sql="select * from elite where catagory='"&catagory&"' and type=1 and location='"&directory&"'"
   'Response.Write sql
   set rs1=conn.Execute(sql)
   do while not rs1.eof 
      dir=trim(rs1("location"))&trim(rs1("title"))&"\"
      disdir=replace(dir,"\","\\")%> 
    <ul>
      <li><a href=# onclick="javascript: givevalue('<%=disdir%>')"><%Response.write trim(rs1("title"))%></a>
         <%ShowSubTree dir,catagory
            %></li>
      </li>
    </ul>
  <%rs1.movenext
   loop
 End Sub 
   catagory=request("catagory")
  if (not session("pass"&catagory)) and not session("admin") then
	Response.Write "<div align=center>对不起，您不能打开本页！</div>"
	Response.End
  end if
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
   
   directory="\"
   sql="select * from elite where catagory='"&catagory&"' and type=1 and location='"&directory&"'"
   'Response.Write sql
   set rs=conn.Execute(sql)%>
 <ul>
  <li><a href=# onclick="javascript: givevalue('\\')">\</a>
  <%do while not rs.eof 
      directory=trim(rs("location"))&trim(rs("title"))&"\"
      disdirectory=replace(directory,"\","\\")%> 
    <ul>
      <li><a href=# onclick="javascript: givevalue('<%=disdirectory%>')"><%Response.write trim(rs("title"))%></a>
      <%ShowSubTree directory,catagory %></li>
      </li>
    </ul>
  <%rs.movenext
   loop%>  
  </li>
</ul>
</body>
</html>
