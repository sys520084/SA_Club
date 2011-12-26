<%function ShowArticle(id,page,UrlTail,ThisTail,ViewtitleAddress,ThisAddress,MemberAddress,MemberBlog,AddAddress,DealAddress,caller)%>
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
<p>
  <%
  set rs=server.CreateObject("ADODB.Recordset")
	sql="update articles set readtimes=readtimes+1 where id="&id
     if page=1 then
        conn.Execute sql
     end if
	 on error goto 0
     sql="select * from articles where id="&id&" order by id"
	 'response.write sql

     rs.Open sql,conn,3
	if rs.EOF then
        Response.Write "<div align='center'>对不起，您要查看的文章没有找到，可能是它已经被删除了。</div>"
        exit function
     end if
	 author=trim(rs("author"))
	 if not author="" then
		temp=split(author,"(")
		author=temp(0)
	 end if  
	 'response.end
     'Response.Write id

     'on error resume next
  %>
</p>
<table width="100%" cellpadding="2">
  <tr> 
    <td class="title1" align="center"> <div style="margin:1em;"><%=trim(rs("title"))%></div></td>
  </tr>
  <tr> 
    <td class="title3" align="center"> 
      <%
				response.write "<a href='"&MemberAddress&"?user="&author&"'>"&rs("author")&"</a> 发表于 <font color='green'><small>"&rs("ondate")&" "&rs("ontime")&"</small></font>"
%>
      
    </td>
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
</table>

<table width="95%" cellpadding="4" bgcolor="#FFFFFF">
	<tr>
	  <td width="50%"><hr size="4" width="90%" align=right></td>
	  <td align=right>&nbsp;</td>
  </tr>
	<tr>
	  <td colspan="2">
	  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<%
		set rs1=server.CreateObject("ADODB.Recordset")
		sql="select * from article_category,usercategory where article_category.category_id=usercategory.id and article_category.article_id="&id
		rs1.Open sql,conn,3
		if not rs1.eof then
		%>
			本文所属的精华目录：<a href="http://www.swarmagents.cn/bs/membership/showelite.asp?user=<%=author%>&cid=<%=rs1("category_id")%>" target="_blank"><%=rs1("category")%></a>
			<font color="#985016">｜ </font>		
		<%end if
		rs1.close
		set rs1=nothing
		tags=rs("tags")
		
		if tags<>"" then
		%>
		本文的标签：
		<%
			tags=replace(tags,"，",",")
			tags=replace(tags," ",",")
			mytags=split(tags,",")
			for i=0 to ubound(mytags)
			if mytags(i)<>"" then%>
			<a href="http://www.swarmagents.cn/bs/forum/viewtitle.asp?tags=<%=mytags(i)%>" target="_blank"><%=mytags(i)%></a>&nbsp;
			<%end if
			next
		end if		
		%>
		</td>
	</tr>
</table>

<table cellpadding="4" width="95%">
	<tr>
		<td width="50%">
			<font color="#985016"><%
				children=rs("children")
				if children<>0 then
					response.write "　　　　评论( <a href='#recommand'>"& children &"</a> )"
				else
					response.write "　　　　评论( 0 )"
				end if

			%> <a href="<%=AddAddress%>?id=<%=id%>&ThisTail=<%=MyUrlEncode(ThisTail&"&page="&page)%>&UrlTail=<%=UrlTail%>"  target="_blank" style="border: 1px solid #999999;padding: 2px;">'发表评论</a> ｜ 阅读(<%=rs("readtimes")%>)
			
			</font>
		</td>
		<td align=right>
		<a href="<%=MemberBlog%>?user=<%=author%>" style="border: 1px solid #999999;padding: 2px;"><%=author%>的blog</a>&nbsp;
			<%
			if session("user")=author or isadmin then%>
			<a href="<%=AddAddress%>?eid=<%=rs("id")%><%if isadmin then%>&admin=ok<%end if%>&ThisTail=<%=MyUrlEncode(ThisTail&"&page="&page)%>&UrlTail=<%=UrlTail%>" style="border: 1px solid #999999;padding: 2px;" target="_blank">'编辑此文</a> &nbsp;
			<a href="<%=DealAddress%>?id=<%=rs("id")%>&action=del&UrlTail=<%=UrlTail%>" onclick="return confirm('确实删除这篇文章么？')" style="border: 1px solid #999999;padding: 2px;" target="_blank">'删除此文</a>&nbsp;
			<a href="../elite/showdir.asp?articleid=<%=id%>" onclick="return js_callpages(this.href)"style="border: 1px solid #999999;padding: 2px;">'加入精华</a>&nbsp;
			<%end if
			if session("user")<>"" then
			%>
			<a href="../elite/favorite.asp?id=<%=id%>" target="_blank" onclick="return js_callpages(this.href)" style="border: 1px solid #999999;padding: 2px;">收藏</a>
			<%end if%>
			
		</td>
	</tr>
</table>
<hr size="4" width="90%">

<%
rs.close
end function


function ShowReplies(id,page,UrlTail,ThisTail,ViewtitleAddress,ThisAddress,MemberAddress,MemberBlog,AddAddress,DealAddress,caller)
sql="select * from articles where fatherid="&id&" order by id"

rs.Open sql,conn,3
if rs.eof then%>
 <table width="98%" border="0" cellpadding="4">
		<tr align="center">
		  <td width="40%" align="left">　　　<a href="<%=ThisAddress%>?<%=ThisTail%>&page=<%=page%>&UrlTail=<%=UrlTail%>" style="border: 1px solid #999999;padding: 2px;">^刷新显示</a></td>
		  <td width="25%"></td>
		  
		  <td width="10%"></td>
		 <td width="10%">
		  </td>
		  <td width="15%"></td>
	  </tr>
</table>
<%
end if
if not rs.eof then
	if page=0 or page="" then
	   page=1
	end if  
	pgSz=100

	rs.PageSize=pgSz
	rs.AbsolutePage=page
	pageCount=rs.PageCount
	rsCount=rs.RecordCount

%>
	<div align="center">所有评论</div>
	<%ShowArticlePageBanner id,page,pageCount,UrlTail,ThisTail,ThisAddress,rsCount%>
	<hr size="1" width="90%">
<%


	j=0
	  do while not rs.EOF and j<rs.PageSize
		 j=j+1
		 author=trim(rs("author"))
		 if not author="" then
		   temp=split(author,"(")
		   author=temp(0)
		   set rs111=server.CreateObject("ADODB.Recordset")
		   sql="select * from memberdes where username='"&author&"'"
		   rs111.open sql,conn,3
		   portrait=rs111("portrait")
		   rs111.close
		   set rs111=nothing
		   pic=showportrait(portrait)&" width=48 height=48"
		   if author=session("user") then
			url="..\membership\uploadimg.asp?portrait="&portrait
			alt="点击更改自己的图标"
		  else
			url=MemberAddress&"?user="&author
			alt="看看"&author
		  end if
		  aaa="<a href="&chr(34)&url&chr(34)&">"
		  enda="</a>"
		 end if%>
		 <a name="recommand"></a>
	<TABLE width="95%" align="center" cellpadding="4" cellspacing="1" bgcolor="#EDECF9">
	  <TR> 
		<TD bgcolor="#FFFFFF"> <table width="100%">
			<tr> 
			  <td rowspan="3" width="90%" valign="top">
			  　　<%=j%>&nbsp;<%=aaa%><img src=<%=pic%> alt="<%=alt%>" border=0> <%=rs("author")%><%=enda%> 于 <font color="green"><small><%=rs("ondate")&" "&rs("ontime")%></small></font> 
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
		<a href="<%=MemberBlog%>?user=<%=author%>" style="border: 1px solid #999999;padding: 2px;"><%=author%>的blog</a>&nbsp;
		 <a href="<%=AddAddress%>?id=<%=rs("id")%>&ThisTail=<%=MyUrlEncode(ThisTail&"&page="&page)%>&UrlTail=<%=UrlTail%>"  target="_blank" style="border: 1px solid #999999;padding: 2px;">'发表评论</a> 
			<%
			s=rsCount mod pgSz
			t=j mod pgSz
			'if t=s and page=pageCount and session("user")=author or isadmin then
			if session("user")=author or isadmin then%>
			<a href="<%=AddAddress%>?eid=<%=rs("id")%><%if isadmin then%>&admin=ok<%end if%>&ThisTail=<%=MyUrlEncode(ThisTail&"&page="&page)%>&UrlTail=<%=UrlTail%>" style="border: 1px solid #999999;padding: 2px;" target="_blank">'编辑此文</a> 
			<a href="<%=DealAddress%>?id=<%=rs("id")%>&action=del&UrlTail=<%=UrlTail%>" onclick="return confirm('确实删除这篇文章么？')" style="border: 1px solid #999999;padding: 2px;" target="_blank">'删除此文</a>&nbsp;&nbsp; 
			<%end if%>
		  </div></td>
	  </tr>
</table>
	<hr size="1" width="90%">
	  <%
	rs.MoveNext
	loop
	rs.close
	ShowArticlePageBanner id,page,pageCount,UrlTail,ThisTail,ThisAddress,rsCount
  end if
end function

function ShowArticlePageBanner(id,page,pageCount,UrlTail,ThisTail,ThisAddress,rsCount)
randomize
formnum=int(rnd()*10)
%>
<form name="frmpage<%=formnum%>" method="post" action="<%=ThisAddress%>?<%=ThisTail%>&UrlTail=<%=UrlTail%>">
	  <input type="hidden" name=id value=<%=id%>>
	  <table width="98%" border="0" cellpadding="4">
		<tr align="center">
		  <td width="40%" align="left">　　　<a href="<%=ThisAddress%>?<%=ThisTail%>&page=<%=page%>&UrlTail=<%=UrlTail%>" style="border: 1px solid #999999;padding: 2px;">^刷新显示</a></td>
		  <td width="25%">第 <font color="red"><%=page%></font> 页/共 <font color="red"><%=pageCount%></font> 页，评论数共 <font color="red"><%=rsCount%></font> 篇</td>
		  
		  <td width="10%"><%if page>=2 then%><a href="<%=ThisAddress%>?<%=ThisTail%>&page=<%=page-1%>&UrlTail=<%=UrlTail%>" style="border: 1px solid #999999;padding: 2px;"><%end if%><<上一页
		  <%if page>=2 then%></a>
		  <%end if%></td>
		 <td width="10%">
		  <%if page<pageCount then%>
		  <a href="<%=ThisAddress%>?<%=ThisTail & "&page=" & (page+1) %>&UrlTail=<%=UrlTail%>" style="border: 1px solid #999999;padding: 2px;"><%end if%>下一页>>
		  <%if page<pageCount then%></a><%end if%></td>
		  <td width="15%">跳转到第 
			<select name="page" onchange="if(this.value!='')frmpage<%=formnum%>.submit();">
			<%for i=1 to pageCount%>
			   <option value=<%=i%> <%if page=i then%> selected <%end if%>><%=i%></option>
			<%next%>   
			</select>
			页</td>
	  </tr>
	</table>
</form>
<%end function%>
