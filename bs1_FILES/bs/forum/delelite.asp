<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>加入精华</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="TEXT/CSS"> 
<!--
body,table {color:#000000;font-family: ??_GB2312; font-size: 9pt; line-height: 12pt}
A:link {text-decoration: none; color:#660000; font-family: "??"; font-size: 9pt; line-height: 12pt}
A:visited {text-decoration: none; color: #660000; font-family: "??"; font-size: 9pt; line-height: 12pt}
A:active {text-decoration: underline; color: #E00000; font-family: "??"; font-size: 9pt; line-height: 12pt}
A:hover {text-decoration: none; color: F00000; font-family: "??"; font-size: 9pt; line-height: 12pt}
-->
</style>
</head>

<body bgcolor="#FFFFFF">
<%
     nowuser=trim(session("user"))
     set conn=server.CreateObject("ADODB.Connection")
     conn.Open OpenString
     set rs=server.CreateObject("ADODB.Recordset")
     catagory=request("catagory")
     sql="select * from power where username='"&nowuser&"' and ((catagory='"&catagory&"' and power=0) or power=2)" 
     rs.Open sql,conn
     isadmin=false
     if rs.EOF then
        Response.Write "<div align=center>对不起，您不能打开本页</div>"
        Response.End 
     end if
     rs.Close 
     id=request("id")
     if id="" then
        Response.write "<div align=center>没有选定文章</div>"
        Response.End
     end if
    sql="select author from articles where id="&id
    set rs=conn.Execute(sql)
    if not rs.eof then
       author=rs(0)
       'Response.Write author
       if author<>"" then
          if instr(1,author,"(")=0 then
             username=author
          else   
             username=left(author,instr(1,author,"(")-1)
          end if
       end if
       'Response.Write username
       'Response.End 
    end if 
    rs.close
    sql="select score from articles where id="&id&" and score=0"
    'Response.Write sql
    set rs=conn.Execute(sql)
    if not rs.EOF then
       Response.Write "<div align=center>这篇文章已经不是精华了！</div>"
       Response.end
    end if
    rs.Close     
    sql="update articles set score=0 where id="&id
    conn.BeginTrans 
    conn.Execute sql
    if not username="" then
       sql="select articles,markscore from users where username='"&username&"'"
       'Response.Write sql
       'Response.End 
       'on error resume next
       rs.Open sql,conn,1,3 
       if not rs.EOF then
         if rs("markscore")>0 then
          rs("markscore")=rs("markscore")-1
         end if
         rs.Update
       end if  
       rs.Close
    end if
    if conn.Errors.Count>0 then
       for i=0 to conn.Errors.Count-1
          Response.Write conn.Errors(i).Description
       next 
       conn.RollbackTrans
       Response.End
    else
       conn.CommitTrans
    end if      
        %>
<div align="center">去除精华成功 </div>
<p><div align="center"><a href="javascript:window.close();">关闭窗口</a></div></p>
</body>
</html>
