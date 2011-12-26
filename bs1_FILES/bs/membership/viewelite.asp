<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
 <!--#include file="../forum/aspFunctions.asp"-->
 <!--#include file="../publicmodule/showarticle.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
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

<script language="javascript">
<!--
function js_callpages(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=330");
  newwin.focus();
  return false;
}
-->
</script>
<%
  UrlTail=request("urltail")
  'response.write UrlTail
  nowuser=trim(session("user"))
  id=0
  id=request("id")
  page=request("page")
  cid=request("cid")
  if page="" then
	page=1
  else
	page=cint(page)
  end if
  ThisTail="id="&id&"&user="&request("user")
  if cid<>"" then
	ThisTail=ThisTail&"&cid="&cid
  end if

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
</head>

<body bgcolor="#FFFFFF">
<div align=left>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="left" valign="top" bgcolor="#DDDDFF" ><!--#include file="../membership/userdisplay.asp"--></td>
    <tr>
      <td>
	   
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>
  <%
  cid=request("cid")
  if cid<>""  then
	sql="select * from usercategory where id="&cid
	set rstemp3=server.CreateObject("ADODB.Recordset")
	rstemp3.open sql,conn,3
	if not rstemp3.eof then
		category=rstemp3("category")
	else
		response.write "<div align=center>对不起，当前目录已经被删除！</div>"
		response.end
	end if
	rstemp3.close
  end if

ShowArticle id,page,UrlTail,ThisTail,"bsblog.asp","viewelite.asp","display.asp","bsblog.asp","../forum/new.asp","../forum/deal.asp","viewelite.asp"
ShowReplies id,page,UrlTail,ThisTail,"bsblog.asp","viewelite.asp","bsblog.asp","bsblog.asp","../forum/new.asp","../forum/deal.asp","viewelite.asp"
%></td>
            <%if cid<>"" then%><td width="200" align="top" valign="top" bgcolor="#DDDDFF">
			<div align="center"><p>&nbsp;</P><p class="bs_blog_title">§<%=category%>目录</p>
	<%
	sql="select * from usercategory where fatherid="&request("cid")
	rstemp3.open sql,conn,3
	do while not rstemp3.eof
	%>
	<div align="center"><a href="showelite.asp?cid=<%=rstemp3("id")%>&user=<%=request("user")%>"><%=rstemp3("category")%></a></div>
	<%
		rstemp3.movenext
	loop
	rstemp3.close

	sql="select top 20 * from article_category,articles where article_category.article_id=articles.id and articles.id<>"&id&" and article_category.category_id="&cid

	'response.write sql
	'response.end

	rstemp3.open sql,conn,3
	do while not rstemp3.eof
	%>
	<p align="center"><a href="viewelite.asp?id=<%=rstemp3("article_id")%>&user=<%=request("user")%>&cid=<%=cid%>"><%=rstemp3("title")%></a></p>
	<%
		rstemp3.movenext
	loop%>
</div>
			</td><%end if%>
          </tr>
        </table> </td>
    </tr>
  </table>
</div>
</body>
</html>
