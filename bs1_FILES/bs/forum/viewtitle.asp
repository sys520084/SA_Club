<!--#include file="../../var.asp"-->
<!--#include file="convert.inc"-->
<!--#include file="aspFunctions.asp"-->
<!--#include file="../publicmodule/showtitles.asp"-->
<!--#include file="../publicmodule/function.asp"-->

<%OpenString=DBBS()%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>察看主题</title>
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
</head>
<body bgcolor="#FFFFFF" text="#CCCCCC" link="#0000CC" vlink="#000000">
<%set conn=server.CreateObject("ADODB.Connection")
  Conn.open OpenString
    'response.write request("subs")

    set rs=server.CreateObject("ADODB.Recordset")
  dim result
  dim resV
  dim sql

  '返回搜索结果
  'response.write request("order")
  result=SearchingTopics("articles")
  resV=split(result,vbTab)
  sql=resV(0)
  urlTail=resV(1)
  isadmin=IsSuperAdmin(Conn,Session("user"))
  page=request("page")
  'response.write page
  order=request("order")
  catagory=request("catagory")
  if page="" then
	page=1
  else
	page=cint(page)
  end if
  'Response.Write sql
   ' Response.End 
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

showPages "viewtitle.asp",UrlTail,"",1
showTitles "viewblog.asp","new.asp","../membership/bsblog.asp",UrlTail,"",rs,"viewtitle.asp"
showPages "viewtitle.asp",UrlTail,"",2%>
  <div>
  <p align=right>
<%if isadmin then%>
      <a href="../admin/admin.asp"><font color="blue">管理</font></a>
<%end if%> </p>
    
  </div>

</body>
</html>
