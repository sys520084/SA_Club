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
      <td align="left" valign="top" bgcolor="#DDDDFF" ><!--#include file="../membership/userdisplay.asp"-->
      </td>
	  <tr>
      <td>
	   
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><%
  dim sql
  if request("user")<>"" and request("cid")<>"" then
	sql="select articles.* from articles,article_category where author='"&request("user")&"' and fatherid=0 and article_category.article_id=articles.id and article_category.category_id="&request("cid")&" order by article_category.article_id desc"
  else
	response.write "<div align=center>打开此页有错误</div>"
	response.end
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
  end if%>
  <%
	sql="select * from usercategory where id="&request("cid")
	set rstemp3=server.CreateObject("ADODB.Recordset")
	rstemp3.open sql,conn,3
	if not rstemp3.eof then
		category=rstemp3("category")
		showDescript=rstemp3("descript")
	else
		response.write "<div align=center>对不起，当前目录已经被删除！</div>"
		response.end
	end if
	rstemp3.close%>
	  <table width="100%" border="0" cellpadding="8" cellspacing="1" bordercolorlight="#FFFFFF" bordercolordark="#000099" bgcolor="#FFFFFF"
  style="TABLE-LAYOUT: fixed; WORD-BREAK: break-all" >
  <tr align="left" bgcolor="white"><td>
  <br>
                <p align=center style="font-size:11pt;"><b><%=category%>目录</B></p>
				<p><%=showDescript%></p>
  </td>
  </tr>
  </table>
  <%
showTitles "viewelite.asp","../forum/new.asp","bsblog.asp","","user="&request("user")&"&cid="&request("cid"),rs,"showelite.asp"
tail="user="&request("user")&"&cid="&request("cid")
if request("reply")<>"" then
	tail=tail&"&replay="&request("reply")
end if
%>
</td>
            <td width="200" bgcolor="#DDDDFF" valign="top"><!--#include file="subcategory.asp"--></td>
          </tr>
        </table> </td>
    </tr>
  </table>
<%showPages "showelite.asp","",tail,1%>
</div>
</body>
</html>
