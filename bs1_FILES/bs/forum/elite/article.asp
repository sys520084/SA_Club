<!--#include file="../../../VAR.asp"-->
<!--#include file="../../publicmodule/function.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:#FFFFFF;}
a:link {
	text-decoration: none;
	color: #00CCFF;
}
a:visited {
	text-decoration: none;
	color: #00CCFF;
}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;
	color: #0099FF;
}
-->
</style>
</head>

<body bgcolor="#003366" text="#FFFFFF" link="#00CCFF" vlink="#00CCFF" alink="#00FFFF">
<%catagory=request("catagory")
  location=request("location")
  id=request("id")
  page=cint(id)
  if location="" then
     location="\"
  end if   
     set conn=server.CreateObject("ADODB.Connection")
     conn.Open OpenString
     set rs=server.CreateObject("ADODB.Recordset")
     sql="select * from articles where id="&id
     'Response.Write sql
     'Response.End 
     rs.Open sql,conn,3
     if rs.EOF then
        Response.Write "<div align=center>�Բ���û���ҵ���ƪ���£�<a href='javascript:history.go(-1)'>�뷵��</a></div>" 
        Response.End
     end if
       t=split(location,"\")
   for i=1 to ubound(t)-1
       for j=1 to i
          add=add&"\"&t(j)
       next
       add=add&"\"
       out=out&"<a href='title.asp?location="&server.URLEncode(add)&"&catagory="&catagory&"'>"&t(i)&"</a>\"
       add=""
   next
   out="<a href='title.asp?location=\&catagory="&catagory&"'>\</a>"&out 
    %>
<div align="center">
  <p><%=catagory%>������-&gt;<%=out%></p>
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#00CCFF" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor="#CCCCFF"> 
      <td bgcolor="#006699"><font color="#33CCFF">���±��⣺<%=trim(rs("title"))%></font></td>
    </tr>
    <tr bgcolor="#FFCCCC"> 
      <td bgcolor="#993366"><font color="#FFCCFF">���ߣ�<%=trim(rs("author"))%>&nbsp;&nbsp;��<%=trim(rs("ondate"))%>ʱ����</font></td>
    </tr>
    <tr> 
      <td><%
       content=rs("content")
       content=ViewImg(content)
       Response.Write content
       %></td>
    </tr>
  </table>
  <hr size="1" color="white">
</div>

</body>
</html>
