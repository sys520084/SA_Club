<%sub showTitles(articleaddress,newaddress,memberaddress,urlTail,ViewTail,rs,caller)%>
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
		 response.end 
         exit sub
    end if%>   

  <table width="100%" border="0" cellpadding="8" cellspacing="1" bordercolorlight="#FFFFFF" bordercolordark="#000099" bgcolor="#FFFFFF"
  style="TABLE-LAYOUT: fixed; WORD-BREAK: break-all" >
    <%total=0
	do while not rs.EOF and total<pgSz
         total=total+1
         id=rs("id")
         mark=""
         if trim(rs("capability"))="0" then
            mark="（空）"
         else 
            mark="("&trim(rs("capability"))&"字节)"   
         end if
         set rs1=server.CreateObject("ADODB.Recordset")
         sql1="select author from articles where fatherid="&id&" order by id desc"
         rs1.Open sql1,conn,3
         replyauthor=""
         if not rs1.EOF then
			replyauthor=rs1(0)
			if right(replyauthor,2)="()" then
				replyauthor=left(replyauthor,len(replyauthor)-2)
			end if
			replyauthor="</font><font color=#6633FF>"&replyauthor&"</font><font color=gray> 于 "
		 end if
		 rs1.close
		 mark1=""
		 if rs("children")<>0 then
			mark1="<font color=gray>(最后回复：" &replyauthor& rs("lastdate") &")</font>"
         end if
		 author=trim(rs("author"))
         if not author="" then
                 temp=split(author,"(")
                 username=temp(0)
         end if  
         if right(author,2)="()" then
			author=left(author,len(author)-2)
		 end if%> 
    <tr align="left" <%if (total mod 2)=1 then%>bgcolor="#f3f6f6"<% else 
	%>bgcolor="#ffffff"<%end if%> > 
	  <td>
	  <%myaddress=articleaddress
	  pic="../img/"&trim(rs("face"))&".gif"

	  if articleaddress="viewblog.asp" then
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
			url=memberaddress&"?user="&username
			alt="看看"&username
		  end if
		  aaa="<a href="&chr(34)&url&chr(34)&">"
		  enda="</a>"
	  end if
	  %>
      <%=aaa%><img src=<%=pic%> align="left" alt="<%=alt%>" border=0><%=enda%>
	  <%if ViewTail<>"" then
			showThisTail="&"&ViewTail
		end if
		if UrlTail<>"" then
			showUrlTail="&"&MyUrlEncode(UrlTail)
		end if%>
	  <b><a href="<%=myaddress%>?id=<%=id&showThisTail&showUrlTail%>" style="font-size:11pt;"><%=trim(rs("title"))%></a></b>
	  <%if not isempty(rs("score")) and rs("score")=1 then%> <font color="red">M</font> <%else%> &nbsp; <%end if%>
	  <small><font color="gray"><%=mark%></font></small>
	  <br>
	  　　- <a href="<%=memberaddress%>?user=<%=username%>"><%=author%></a>
      <small><font color="green">(<%=trim(rs("ondate"))%>)</font></small>
	  <br>
	  <font color="#444444"><%=GetAbstract(rs("content"),300)&"..."%></font>
	  <div align="right">
	  <hr align="right" width="400" size="1">
		<font color="#985016">评论( <a href="<%=myaddress%>?id=<%=id%>&UrlTail=<%=MyUrlEncode(UrlTail)%>"><%=rs("children")%></a> ) <a href="<%=newaddress%>?id=<%=id%>&UrlTail=<%=MyUrlEncode(UrlTail)%>" target="_Blank">发表新评论</a> | 阅读(<%=rs("readtimes")%>) </font>
		<%=mark1%>
		<%if session("user")=username then
			if articleaddress="viewelite.asp" and request("cid")<>"" then
				showdirtail="articleid="&id&"&categoryid="&request("cid")
				alt="移动到其它精华目录"
			else
				showdirtail="articleid="&id
				alt="加入精华目录"
			end if
		%>
		|<a href="../elite/showdir.asp?<%=showdirtail%>" target="_blank" onclick="return js_callpages(this.href)"><img src="../images/elite.jpg" alt=<%=alt%> width="14" height="14" border="0">精华</a> <%if articleaddress="viewelite.asp" and request("cid")<>"" then%><a href="../elite/delelite.asp?<%=showdirtail%>" target="_blank" onclick="return js_callpages(this.href)"><img src="../images/delete.gif" alt=从当前目录删除这篇文章 width="14" height="14" border="0"></a> <%end if%>
		<%else
		 if session("user")<>"" then%>
		|<%if caller<>"showfavorite.asp" then%><a href="../elite/favorite.asp?id=<%=id%>" target="_blank" onclick="return js_callpages(this.href)"><img src="../images/elite.jpg" alt=收藏这篇文章 width="14" height="14" border="0">收藏</a><%end if%> <%if caller="showfavorite.asp" then%><a href="../elite/delfavorite.asp?id=<%=id%>" target="_blank" onclick="return js_callpages(this.href)"><img src="../images/delete.gif" alt=不收藏这篇文章 width="14" height="14" border="0">去收藏</a> <%end if%>
		
      <% end if
	  end if%>
	  <%if isadmin and caller="viewtitle.asp" then%>|<a href="../elite/putfirst.asp?id=<%=id%>" target="_blank"><%if rs("score")=1 then%>下<%else%>上<%end if%>首页</a><%end if%>
	  </div>
	  </td>
  </tr>
    <%rs.MoveNext
          loop%> 
  </table>
<%end sub
sub showPages(viewAddress,urlTail,tail,times)
	if urlTail<>"" then
		urlTail=MyHtmlEncode(UrlTail)
		if left(urlTail,1)<>"&" then
			urlTail="&"&UrlTail
		end if
	end if
%>
 
  <table width="100%" border="0">
    <tr align="center"> 
	<form name="form1" action="<%=viewAddress%>" method="post" onsubmit="checkout();">
      <td width="40%" height="30" align="left"><a href="<%=viewAddress%>?page=<%=page%><%if tail<>"" then response.write("&"&tail) end if%><%=urlTail%>" style="border: 1px solid #999999;padding: 2px;">^刷新显示</a>&nbsp;&nbsp;
<% if viewAddress<>"showelite.asp" and viewAddress<>"showfavorite.asp" then%>
<input name="searchvalue" type="text" value="" size="10">
<select name="item">
  <option value="title" selected>题目</option>
  <option value="author">作者</option>
  <option value="tags">标签</option>
  <option value="content">内容</option>
</select>
<input type="hidden" name="shortsearch" value="ok">
<%if viewAddress="show-all-elites.asp" then%>
<input type="hidden" name="score" value="1">
<%end if%>
<input type="submit" name="finder" value="搜索">
<input type="button" name="finder2" value="高级" onclick="javascript:window.location.replace('../../bs/forum/finder.asp')">
<%end if%>
	  </td>
	  </form>	 
      <td width="25%" height="24">第 <font color=red><%=page%></font> 页/共 <font color=red><%=pgCount%></font> 页，文章数共 <font color=red><%=rsCount%></font> 篇</td>
      <td width="9%" height="24"> 
        <%if page>1 then%>
        <a href="<%=viewAddress%>?page=<%=page-1%><%if tail<>"" then response.write("&"&tail) end if%><%=urlTail%>" style="border: 1px solid #999999;padding: 2px;"> 
        <%end if%>
        <<上一页</a></td>
      <td width="7%" height="24"> 
        <%if page<pgCount then%>
        <a href="<%=viewAddress%>?page=<%=page+1%><%if tail<>"" then response.write("&"&tail) end if%><%=urlTail%>" style="border: 1px solid #999999;padding: 2px;"> 
        <%end if%>
        下一页>></a></td>
      
	  <form method="post" action="<%=viewAddress%>?<%if tail<>"" then response.write(tail) end if%><%=urlTail%>" name="frm<%=times%>" onsubmit="if(frm<%=times%>.page.value!='no'){frm<%=times%>.submit();}">
	  <td width="16%" height="24">
	  跳转到第 
        <input type="hidden" name="flag" value="ok">
        <select name="page" onchange="frm<%=times%>.submit();">
          <%for i=1 to pgCount%> 
          <option value=<%=i%> <%if page=i then%> selected <%end if%>><%=i%></option>
          <%next%> 
        </select>
        页
		</td>
	  </form>
    </tr>
  </table>

<%end sub%>