<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<!--#include file="convert.inc"-->
<!--#include file="aspFunctions.asp"-->
<!--#include file="../publicmodule/showArticle.asp"-->

<%
'response.write request("id")
OpenString=DBBS()%>
<html>
<head>
<%

  UrlTail=request("urltail")
  'response.write UrlTail
  nowuser=trim(session("user"))
  id=0
  id=request("id")
  page=request("page")
  if page="" then
	page=1
  else
	page=cint(page)
  end if
  ThisTail="id="&id

  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  sql="select * from articles where id="&id
  set rs=server.CreateObject("ADODB.Recordset")
	 rs.open sql,conn,3
	 if rs.eof then
		response.write "要查的文章不存在"
		response.end
	end if
  isadmin=IsSuperAdmin(Conn,Session("user"))

%>
<title><%=rs("title")%></title>
<%rs.close%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/define.css" rel="stylesheet" type="text/css" />

<style type="text/css">
<!--
.content {
	font-family: "宋体";
	font-size: 10pt;
	font-style: normal;
	line-height: 20px;
}
-->
a:link {text-decoration: none; ; font-size: 9pt; color: #0066FF}
a:visited {text-decoration: none;; color: #0033CC}
a:hover {text-decoration: underline; color: #0033CC}
a:active {text-decoration:none;color: #0066FF}
</style>
</head>

<body>
<A name="top"></A>

<br>
<p class="title3">　<a href="viewtitle.asp?<%=MyUrlDecode(UrlTail)%>" style="border: 1px solid #999999;padding: 2px;">←返回文章列表</a> <a href="viewblog.asp?<%="id="&id%>&page=<%=page%>&UrlTail=<%=UrlTail%>" style="border: 1px solid #999999;padding: 2px;">^刷新显示</a></p>

<%
ShowArticle id,page,UrlTail,ThisTail,"viewtitle.asp","viewblog.asp","../membership/display.asp","../membership/bsblog.asp","new.asp","deal.asp","viewblog.asp"
ShowReplies id,page,UrlTail,ThisTail,"viewtitle.asp","viewblog.asp","../membership/bsblog.asp","../membership/bsblog.asp","new.asp","deal.asp","viewblog.asp"
%>
</p>
</body>
</html>
