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
  sql="select * from files order by id desc"
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
%>
<script language="javascript">
<!--
function js_callpages(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=530");
  newwin.focus();
  return false;
}
-->
</script>
  <hr size="1">
<%  if rs.EOF then
         Response.Write "<div align='center'>对不起，没有找到任何文章。</div>"
         Response.end
    end if%>   

  <table width="100%" border="0" cellpadding="8" cellspacing="1" bordercolorlight="#FFFFFF" bordercolordark="#000099" bgcolor="#FFFFFF"
  style="TABLE-LAYOUT: fixed; WORD-BREAK: break-all" >
    <%total=0
	do while not rs.EOF and total<pgSz
         total=total+1
         id=rs("id")
         
		 author=trim(rs("username"))
         %> 
    <tr align="left" <%if (total mod 2)=1 then%>bgcolor="#f3f6f6"<% else 
	%>bgcolor="#ffffff"<%end if%> > 
	  <td>
	  <%myaddress=articleaddress

	  
		  set rs111=server.CreateObject("ADODB.Recordset")
		  sql="select * from memberdes where username='"&author&"'"
		  rs111.open sql,conn,3
		  portrait=rs111("portrait")
		  rs111.close
		  set rs111=nothing
		  'response.write portrait
		  pic=showportrait(portrait)&" width=48 height=48"
		  if author=session("user") then
			url="..\membership\uploadimg.asp?portrait="&portrait
			alt="点击更改自己的图标"
		  else
			url=memberaddress&"?user="&username
			alt="看看"&username
		  end if
		  aaa="<a href="&chr(34)&url&chr(34)&">"
		  enda="</a>"
	  %>
      <%=aaa%><img src=<%=pic%> align="left" alt="<%=alt%>" border=0><%=enda%>
	  <%if ViewTail<>"" then
			showThisTail="&"&ViewTail
		end if
		if UrlTail<>"" then
			showUrlTail="&"&MyUrlEncode(UrlTail)
		end if%>
	  <b><a href="../upload/download.asp?id=<%=rs("id")%>" style="font-size:11pt;"><%=trim(rs("title"))%></a> (<%=trim(rs("bytes")/1000)%>K)</b>
	  &nbsp;
	  <small><font color="gray"><%=mark%></font></small>
	  <br>
	  　　- <a href="<%=memberaddress%>?user=<%=username%>"><%=author%></a>
      <small><font color="green">(<%=trim(rs("uploaddate"))%>)</font></small>
	  <br>
	  <font color="#444444"><%=GetAbstract(rs("des"),300)&"..."%>
	  <%if not isnull(rs("author")) and rs("author")<>"" then
	  	response.write "<br>作者："&rs("author")
		end if
		if not isnull(rs("keywords")) and rs("keywords")<>"" then
	  	response.write "<br>关键字："&rs("keywords")
		end if
		if not isnull(rs("journal")) and rs("journal")<>"" then
	  	response.write "<br>发表刊物："&rs("journal")&"&nbsp;"&rs("publishtime")
		end if
		
		%>
	  </font>
	  <hr size="1" width="400">
		<font color="#985016">下载(<%=rs("times")%>) | <%if isadmin then%>
      <a href="../upload/delete.asp?id=<%=rs("id")%>"><font color="blue">删除</font></a> | <a href="../upload/move.asp?id=<%=rs("id")%>">移动到资源下载</a>
<%end if%></font>
		<%=mark1%>
		<%if session("user")=username then
			if articleaddress="viewelite.asp" and request("cid")<>"" then
				showdirtail="articleid="&id&"&categoryid="&request("cid")
				alt="移动到其它精华目录"
			else
				showdirtail="articleid="&id
				alt="加入精华目录"
			end if
		%><a href="../elite/showdir.asp?<%=showdirtail%>" target="_blank" onClick="return js_callpages(this.href)"><img src="../images/elite.jpg" alt=<%=alt%> width="14" height="14" border="0"></a> <%if articleaddress="viewelite.asp" and request("cid")<>"" then%><a href="../elite/delelite.asp?<%=showdirtail%>" target="_blank" onClick="return js_callpages(this.href)"><img src="../images/delete.gif" alt=从当前目录删除这篇文章 width="14" height="14" border="0"></a> <%end if%>
      <%end if%>
		</td>
  </tr>
    <%rs.MoveNext
          loop%> 
  </table>
   <form method="post" action="viewfiles.asp" name="frm" onSubmit="if(frm.page.value!='no'){frm.submit();}">
  <table width="100%" border="0">
    <tr align="center"> 
      <td width="40%" height="24" align="left"><a href="viewfiles.asp?page=<%=page%>" style="border: 1px solid #999999;padding: 2px;">^刷新显示</a></td>
      <td width="25%" height="24">第 <font color=red><%=page%></font> 页/共 <font color=red><%=pgCount%></font> 页，文章数共 <font color=red><%=rsCount%></font> 篇</td>
      <td width="9%" height="24"> 
        <%if page>1 then%>
        <a href="viewfiles.asp?page=<%=page-1%>" style="border: 1px solid #999999;padding: 2px;"> 
        <%end if%>
        <<上一页</a></td>
      <td width="7%" height="24"> 
        <%if page<pgCount then%>
        <a href="viewfiles.asp?page=<%=page+1%>" style="border: 1px solid #999999;padding: 2px;"> 
        <%end if%>
        下一页>></a></td>
      <td width="16%" height="24">跳转到第
        <input type="hidden" name="flag" value="ok">
        <select name="page" onChange="frm.submit();">
          <%for i=1 to pgCount%> 
          <option value=<%=i%> <%if page=i then%> selected <%end if%>><%=i%></option>
          <%next%> 
        </select>
        页</td>
    </tr>
  </table>
  </form>
  <div>
  <p align=right>
 </p>
    <div align="center"> 
      <p><font color="#333333">集智俱乐部·版权所有 <font face="Arial, Helvetica, sans-serif"><br>
        Copyright 2007-2027 Clustering Intelligence Club</font></font></p>
    </div>
  </div>

</body>
</html>
