<%response.redirect("http://www.swarmagents.cn/bs/forum/viewarticle.asp?id="&id)%>
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
<%
  set rs=server.CreateObject("ADODB.Recordset")
	sql="update articles set readtimes=readtimes+1 where id="&id
     if page=1 then
        conn.Execute sql
     end if
	 on error goto 0
     sql="select * from articles where id="&id&" or fatherid="&id&" order by id"
     pgSz=10
	 'response.write sql

     rs.Open sql,conn,3
	if rs.EOF then
        Response.Write "<div align='center'>对不起，您要查看的文章没有找到，可能是它已经被删除了。</div>"
        Response.End
     end if
	 'response.end
     'Response.Write page
     if page=0 then
        page=1
     end if   
     rs.PageSize=pgSz
     rs.AbsolutePage=page
     pageCount=rs.PageCount
     rsCount=rs.RecordCount

     'on error resume next
  %>
	<form name="frmpage" method="post" action="viewarticle.asp?<%=ThisTail%>&UrlTail=<%=UrlTail%>">
	  <input type="hidden" name=id value=<%=id%>>
	  <table width="98%" border="0" cellpadding="4">
		<tr align="center">
		  <td width="40%" align="left"><a href="viewtitle.asp?<%=MyUrlDecode(UrlTail)%>" style="border: 1px solid #999999;padding: 2px;">←返回文章列表</a>　<a href="viewarticle.asp?<%=ThisTail%>&page=<%=page%>&UrlTail=<%=UrlTail%>" style="border: 1px solid #999999;padding: 2px;">^刷新显示</a></td>
		  <td width="25%">第 <font color="red"><%=page%></font> 页/共 <font color="red"><%=pageCount%></font> 页，评论数共 <font color="red"><%=rsCount%></font> 篇</td>
		  
		  <td width="10%"><%if page>=2 then%><a href="viewarticle.asp?<%=ThisTail%>&page=<%=page-1%>&UrlTail=<%=UrlTail%>#recommand" style="border: 1px solid #999999;padding: 2px;"><%end if%><<上一页
		  <%if page>=2 then%></a>
		  <%end if%></td>
		 <td width="10%">
		  <%if page<pageCount then%>
		  <a href="viewarticle.asp?<%=ThisTail & "&page=" & (page+1) %>&UrlTail=<%=UrlTail%>#recommand" style="border: 1px solid #999999;padding: 2px;"><%end if%>下一页>>
		  <%if page<pageCount then%></a><%end if%></td>
		  <td width="15%">跳转到第 
			<select name="page" onchange="if(this.value!='')frmpage.submit();">
			<%for i=1 to pageCount%>
			   <option value=<%=i%> <%if page=i then%> selected <%end if%>><%=i%></option>
			<%next%>   
			</select>
			页</td>
	  </tr>
	</table>
	</form>
	<hr size="1" width="90%">
<%j=0
  do while not rs.EOF and j<rs.PageSize
     j=j+1
	 author=trim(rs("author"))
     if not author="" then
       temp=split(author,"(")
       username=temp(0)
     end if%> 
	<TABLE cellpadding="4" cellspacing="1" width="95%" align="center">
	  <TR> 
		<TD bgcolor="#FFFFFF"> <table width="100%">
			<tr> 
			  <td rowspan="3" width="90%" valign="top">
			  　　　<%="<img src='../img/"&rs("face")&".gif'>"%> <a href="../membership/display.asp?name=<%=author%>" onclick="return js_callpages(this.href)"><%=rs("author")%></a> 于 <font color="green"><small><%=rs("ondate")&" "&rs("ontime")%></small></font> 
			   <font color="blue"><b><%=rs("title")%></b></font>
			   </td>
			  <td width="10%" align="right"> <a href="#top">△TOP</a> </td>
			</tr>
		  </table></td>
	  </tr>
	  <tr> 
		<td> <div class="text1"> 
			<%
				content=rs("content")
				content=ViewImg(content)
				response.write content
				%>
		  </div></td>
	  </tr>
	  <tr>
		<td> <div align="right">
		 <a href="new.asp?id=<%=id%>&ThisTail=<%=MyUrlEncode(ThisTail&"&page="&page)%>&UrlTail=<%=UrlTail%>"  target="_blank" style="border: 1px solid #999999;padding: 2px;">'回复此文</a> 
			<%
			s=rsCount mod pgSz
			t=j mod pgSz
			'response.write session("user") & "," & isadmin
			if t=s and page=pageCount and session("user")=username or isadmin then%>
			<a href="new.asp?eid=<%=rs("id")%><%if isadmin then%>&admin=ok<%end if%>&ThisTail=<%=MyUrlEncode(ThisTail&"&page="&page)%>&UrlTail=<%=UrlTail%>" style="border: 1px solid #999999;padding: 2px;">'编辑此文</a> 
			<a href="deal.asp?id=<%=rs("id")%>&action=del&UrlTail=<%=UrlTail%>" onclick="return confirm('确实删除这篇文章么？')" style="border: 1px solid #999999;padding: 2px;">'删除此文</a>&nbsp;&nbsp; 
			<%end if%>
		  </div></td>
	  </tr>
	</table>
	<hr size="1" width="90%">
<br>
<%rs.MoveNext
  'Response.Write j
  loop
  rs.Close
  set rs=nothing%>
	<form name="frmpage" method="post" action="viewarticle.asp?<%=ThisTail%>&UrlTail=<%=UrlTail%>">
	  <input type="hidden" name=id value=<%=id%>>
	  <table width="98%" border="0" cellpadding="4">
		<tr align="center">
		  <td width="40%" align="left">　　　<a href="viewarticle.asp?<%=ThisTail%>&page=<%=page%>&UrlTail=<%=UrlTail%>" style="border: 1px solid #999999;padding: 2px;">^刷新显示</a></td>
		  <td width="25%">第 <font color="red"><%=page%></font> 页/共 <font color="red"><%=pageCount%></font> 页，评论数共 <font color="red"><%=rsCount%></font> 篇</td>
		  
		  <td width="10%"><%if page>=2 then%><a href="viewarticle.asp?<%=ThisTail%>&page=<%=page-1%>&UrlTail=<%=UrlTail%>#recommand" style="border: 1px solid #999999;padding: 2px;"><%end if%><<上一页
		  <%if page>=2 then%></a>
		  <%end if%></td>
		 <td width="10%">
		  <%if page<pageCount then%>
		  <a href="viewarticle.asp?<%=ThisTail & "&page=" & (page+1) %>&UrlTail=<%=UrlTail%>#recommand" style="border: 1px solid #999999;padding: 2px;"><%end if%>下一页>>
		  <%if page<pageCount then%></a><%end if%></td>
		  <td width="15%">跳转到第 
			<select name="page" onchange="if(this.value!='')frmpage.submit();">
			<%for i=1 to pageCount%>
			   <option value=<%=i%> <%if page=i then%> selected <%end if%>><%=i%></option>
			<%next%>   
			</select>
			页</td>
	  </tr>
	</table>
	</form>
</body>
</html>
