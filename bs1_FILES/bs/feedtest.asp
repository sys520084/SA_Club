<!--#include file="../VAR.asp"-->
<!--#include file="publicmodule/function.asp"-->
<!--#include file="forum/aspFunctions.asp"-->
<%
Response.ContentType = "text/XML"
Dim sql,rs,dd,conn
i=1
%>
<?xml version="1.0" encoding="gb2312" ?>
<bookinfo>
<%Do While i<10%>
    <book>
       <title><%="title"%></title>
       <author><%="author"%></author>
    <publish>
             <publisher><%="publisher"%></publisher>
       <ISBN><%="ISBN"%></ISBN>
    </publish>
    <price><%="price"%></price>
    </book>
<%
i=i+1
Loop
%>
</bookinfo>