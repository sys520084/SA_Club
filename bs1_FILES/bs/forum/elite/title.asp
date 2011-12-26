<!--#include file="../../../VAR.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>精华区</title>
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
	color: #00CCFF;
}
-->
</style>
</head>

<body bgcolor="#003366">
<%catagory=request("catagory")
  location=request("location")
  page=request("page")
  if page="" then
     page=1
  end if      
  page=cint(page)
  if location="" then
     location="\"
  end if   
  if catagory<>"" then
     set conn=server.CreateObject("ADODB.Connection")
     conn.Open OpenString
     set rs=server.CreateObject("ADODB.Recordset")
     sql="select * from elite where location='"&location&"' and catagory='"&catagory&"' order by type desc,title"
    ' Response.Write sql
     rs.Open sql,conn,3
     if rs.EOF then
        if location="\" then
           Response.Write "<div align=center>对不起，这个城市没有整理的精华区！<a href='javascript:history.go(-1)'>请返回</a></div>" 
        else
           Response.Write "<div align=center>对不起，这个目录里面尚没有精华文章！<a href='javascript:history.go(-1)'>请返回</a></div>"
        end if
        Response.End
     end if
     pgSz=20
     rs.PageSize=pgSz
     rs.AbsolutePage=page
     rsCount=rs.RecordCount
     pgCount=rs.PageCount
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
   j=0  %>
<div align="center">
  <p><%=catagory%>精华区-&gt;<%=out%></p>
  <table width="100%" border="0">
    <tr align="center" bgcolor="#000033"> 
      <td width="7%"><font color="#FFFFCC">类型</font></td>
      <td width="56%"><font color="#FFFFCC">题目</font></td>
      <td width="14%"><font color="#FFFFCC">作者</font></td>
      <td width="12%"><font color="#FFFFCC">整理人</font></td>
      <td width="11%"><font color="#FFFFCC">整理日期</font></td>
    </tr>
    <%j=0
     do while not rs.EOF and j<pgSz
         j=j+1
         if (j mod 2)=1 then%> 
    <tr bgcolor="#336699"> 
       <%else%> 
    <tr bgcolor="#253489">
       <%end if%> 
      <td width="7%"> <%if rs("type")=0 then%> 文章 <%else%> 文件夹 <%end if%> </td>
      <td width="56%"> <%if rs("type")=1 then%> <a href="title.asp?catagory=<%=catagory%>&location=<%=trim(rs("location"))&trim(rs("title"))&"\"%>"> 
        <%=trim(rs("title"))%></a> <%else%> 
        <a href="article.asp?id=<%=rs("article_id")%>&catagory=<%=catagory%>&location=<%=location%>"><%=trim(rs("title"))%></a> <%end if%> 
      </td>
      <td width="14%"><%if rs("type")=1 then%>&nbsp;<%else%><%=rs("author")%><%end if%></td>
      <td width="12%"><%=trim(rs("deal_man"))%></td>
      <td width="11%"><%=rs("deal_date")%></td>
    </tr>
    <%rs.Movenext
    loop%> 
  </table>
  <hr>
  <form name="frmPage" method="get" action="title.asp">
    <table width="100%" border="0">
      <tr align="center" bgcolor="#FFCCCC"> 
        <td bgcolor="#003366">分页： 
          <input type=hidden name="catagory" value="<%=catagory%>">
          <input type=hidden name="location" value=<%=location%>>
          <%if page>1 then%> <a href="title.asp?page=<%=page-1%>&catagory=<%=catagory%>&location=<%=location%>">上一页</a> 
          &nbsp; <%end if
      if page<pgCount then%> <a href="title.asp?page=<%=page+1%>&catagory=<%=catagory%>&location=<%=location%>">下一页</a> 
          &nbsp; <%end if%>跳到： 
          <select name="page" onChange="frmPage.submit();">
            <% for i=1 to pgCount%> 
            <option value=<%=i%><%if i=page then%> selected<%end if%>><%=i%></option>
            <% next %> 
          </select>
          页，共<%=rsCount%>篇精华 </td>
      </tr>
    </table>
    </form>
  </div>
<%end if%>
</body>
</html>
