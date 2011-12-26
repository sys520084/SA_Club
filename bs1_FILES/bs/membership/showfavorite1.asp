<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
 <!--#include file="../forum/aspFunctions.asp"-->
 <!--#include file="../publicmodule/showtitles.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/define.css" rel="stylesheet" type="text/css" />
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
<body bgcolor="#FFFFFF">
<div align=left>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="left" valign="top" bgcolor="#DDDDFF" ><!--#include file="userdisplay.asp"-->
      </td>
	  <tr>
      <td>
	   <%
  dim sql
  if request("user")<>"" then
	sql="select articles.id as id,* from articles,favorite where favorite.username='"&request("user")&"' and fatherid=0 and articles.id=favorite.article_id order by favorite.ondate desc"
  end if
  isadmin=IsSuperAdmin(Conn,Session("user"))
  page=request("page")
  'response.write page
  if page="" then
	page=1
  else
	page=cint(page)
  end if
  'Response.Write sql
  'Response.End 
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
tail="user="&request("user")
if request("reply")<>"" then
	tail=tail&"&replay="&request("reply")
end if
showTitles "viewelite.asp","../forum/new.asp","bsblog.asp","",tail,rs,"showfavorite.asp"
showPages "showfavorite.asp","",tail,1%>
	  </td>
    </tr>
  </table>
  <p>&nbsp;</p>
</div>
</body>
</html>
