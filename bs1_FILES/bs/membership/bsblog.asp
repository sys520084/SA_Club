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
div,td,p,li{font-size:9pt; line-height:14pt; font-family:"����";color:#333333;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0066FF}
a:visited {text-decoration: none;; color: #0033CC}
a:hover {text-decoration: underline; color: #0033CC}
a:active {text-decoration;color: #0066FF}
.art {
	font-size: 10pt;
	color: #0000CC;
	font-style: normal;
}
.style1 {color: #FF0000}
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
<p></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="left" valign="top" bgcolor="#DDDDFF" ><!--#include file="userdisplay.asp"-->
      </td>
    <tr>
      <td>
	    <p>&nbsp;</p>
	    <table width="90%" border="0" align="center" cellpadding="2" cellspacing="0" bordercolor="#6A6AFF">
        <tr>
          <td align="center"><span class="blogname"><%=show%>�ľ������£�</span></td>
        </tr>
        <tr>
          <td>
		  <%  isadmin=IsSuperAdmin(Conn,Session("user"))%>

		  <%


  dim sql
  if request("user")<>"" then
	sql="select top 5 articles.id as id,article_category.id as cid,* from articles,article_category where articles.author='"&request("user")&"' and articles.fatherid=0 and articles.id=article_category.article_id order by articles.lastdate desc"
	tail="user="&request("user")
  else
  	response.write "û��ѡ��鿴�û�"
	response.end 
  end if
  'Response.Write sql
  'Response.End 
  pgSz=5
  rs.Open sql,conn,3
  'Response.End
  if rs.eof then
  %><div align="center">�ޣ�</div>
<%
  else
  	showTitles "viewelite.asp","../forum/new.asp","bsblog.asp","",tail,rs,"blogfirst.asp"%>
		<div align=left><a href="oldbsblog.asp?user=<%=request("user")%>&elite=1">����...</a></div><%
  end if
  rs.close
%>

		  </td>
        </tr>
		</table>
		<p>&nbsp;</p>
		<table width="90%" border="0" align="center" cellpadding="2" cellspacing="0" bordercolor="#6A6AFF">
        <tr>
          <td align="center"><p><span class="blogname"><%=show%>��������£�</span> </p>            </td>
        </tr>
        <tr>
          <td><p></p>
            <p>
              <%
	sql="select top 5 * from articles where author='"&request("user")&"' and fatherid=0 order by lastdate desc"
  'Response.Write sql
  'Response.End 
  rs.Open sql,conn,3
  'Response.End
  if rs.eof and request("open")=1 then
  	response.redirect "../forum/viewtitle.asp?order=lastdate+desc"
  %>
            </p>
            <div align="center">�ޣ�</div>
            <%
  else
  	'response.write "ok"
  	showTitles "viewelite.asp","../forum/new.asp","bsblog.asp","",tail,rs,"blogfirst.asp"
	%>
	<div align=left><a href="oldbsblog.asp?user=<%=request("user")%>">����...</a></div>
	<%
  end if
  rs.close
%></td>
        </tr>
		</table>
		<p>&nbsp;</p>
		<table width="90%" border="0" align="center" cellpadding="2" cellspacing="2" bordercolor="#6A6AFF">
        
		  <%
	sql="select top 5 favorite.ondate as ondate, articles.ondate as ddate,articles.id as id,favorite.id as fid,* from articles,favorite where articles.author='"&request("user")&"' and articles.id=favorite.article_id and favorite.username<>'"&request("user")&"' order by favorite.id desc"
  'Response.Write sql
  'Response.End 
  rs.Open sql,conn,3
  if not rs.eof then%>
  <tr>
          <td bgcolor="#FEFDE2">
		  <p align="center"><span class="title"><%=show%>�����±��ղص������</span></p>
	    <ul>
		<%do while not rs.eof%>
	      <li><%=rs("username")%>��<%=rs("ondate")%>�ղ���<%=show%>�����£�<a href=""><%=rs("title")%></a>����ǩΪ��<%=rs("tag")%></li>
		  <%rs.movenext
		  loop%>
        </ul>
		</td>
        </tr>
  <%end if
  rs.close%>
        <%sql="select top 5 favorite.ondate as ondate, articles.ondate as ddate,articles.id as id,favorite.id as fid,* from articles,favorite where favorite.username='"&request("user")&"' and articles.id=favorite.article_id order by favorite.id desc"
  'Response.Write sql
  'Response.End 
  rs.Open sql,conn,3
  if not rs.eof then
  %><tr>
          <td height="188" bgcolor="#F2F0FF">
          <p align="center"><span class="title"><%=show%>�ղر������µ������</span></p>
	    <ul>
		<%do while not rs.eof%>
	      <li><%=show%>��<%=rs("ondate")%>�ղ���<a href="bsblog.asp?user=<%=rs("author")%>"><%=rs("author")%></a>�����£�<a href=""><%=rs("title")%></a>����ǩΪ��<%=rs("tag")%></li>
		  <%rs.movenext
		  loop%>
        </ul>
		</td>
        </tr>
        <%
  end if
  rs.close
  %>
      </table>      
        <p></p>
 <%
 if session("user")=request("user") then
  detailcontent="<P><STRONG>��ϲ������</STRONG>��</P><P><STRONG>��ϲ��������</STRONG>��</P><P><STRONG>��ϲ��������</STRONG>��</P><P><STRONG>��ϲ���ĵط�</STRONG>��</P><P><STRONG>��ϲ�����˶�</STRONG>��"
  sql="select * from memberdes where username='"&session("user")&"'"
  rs.open sql,conn,3
  if not rs.eof then
  	detail=rs("detailcontent")
	'detail=MyhtmlEncode(detail)
	if detail<>"" then
		'response.write "ol"
		detail=replace(detail," ","")
		detail=replace(detail,chr(9),"")
		detail=replace(detail,chr(10),"")
		detail=replace(detail,chr(13),"")
	end if
	detail=left(detail,len(detailcontent))
	'response.write MyhtmlEncode(detailcontent)
	'response.write detail
  	if trim(detail)=detailcontent or detail="" then
		 %>
        <div align="center"><%=show%><span class="style1">�ĸ������ϻ�û����ɣ���Ͽ�<a href="persenal.asp">���</a></span>��</div>  
		<%
	end if
   end if      
   rs.close
  end if%>
      <p>&nbsp;</p>        <p>&nbsp;</p></td>
    </tr>
  </table>
  <p>&nbsp;</p>
</div>
</body>
</html>
