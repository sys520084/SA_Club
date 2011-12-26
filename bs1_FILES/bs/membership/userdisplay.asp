	<%
  username=request("user")

  if trim(username)="" then
     Response.Write "<div align =center>对不起，您没有选中要查看的用户。"
	 response.write "<a href='relogin.asp?UrlTail=../forum.asp' target='_Parent'>请登陆</a></div>"
     Response.End
  end if
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  set rs=server.CreateObject("ADODB.Recordset")
  sql="select * from memberdes,detail where memberdes.username='"&username&"' and memberdes.username=detail.username"
  rs.Open sql,conn,3
  if not rs.eof then
	portrait=rs("portrait")
	showtitle=rs("showtitle")
	story=rs("readme")
	email=rs("email")
  end if
  alt="给他(她)发消息"
  link="../message/sendmsg.asp?name1="&trim(username)
  if session("user")=username then
	  	alt="点击更改头像"
		link="uploadimg.asp?portrait="&portrait
  end if
  rs.close
       'on error goto 0
	  %> 
<table width="100%" height="153" border="0" cellpadding="0" cellspacing="0" background="../images/bg_bs.gif">
  <tr>
    <td width="153" rowspan="4"><img src="../images/bs_tree.gif" alt="<%=trim(showtitle)%>" width="153" height="153" /></td>
    <td height="35" colspan="2">　　<em><span class="bs_blog_title"><%=trim(showtitle)%></span></em></td>
  </tr>
  <tr>
    <td height="24" colspan="2">　　　　<span class="bs_blog_intro">『<%
	if len(trim(story))=0 or isnull(story) then 
		response.write "... ... ... ..." 
	end if
	response.write trim(story)
	%>』<span> <%if session("user")=username then%><a href="persenal.asp?user=<%=username%>"><img src="../images/edit.gif" width="20" height="20" alt="修改用户信息" border="0">
	  <%end if%>
	  </td>
  </tr>
  <tr>
    <td width="89" rowspan="2" align="center" valign="middle"><img src="<%=ShowPortrait(portrait)%>" alt="<%=alt%>" width=48 height=48 border="0"   bordercolor="#000000"></td>
    <td width="781" height="30">&nbsp;</td>
  </tr>
  <tr>
   <%
  username=request("user")
	show=trim(username)
  if username=session("user") then
  	show="我"
  end if
  set rstemp2=server.CreateObject("ADODB.Recordset")
  sql="select * from usercategory where username='"&username&"' and fatherid=-1"
  rstemp2.open sql,conn,3
  showbanner=""
  tail=""
  do while not rstemp2.eof
	cname=rstemp2("category")
	cid=rstemp2("id")
	showbanner=showbanner&"<a href='showelite.asp?user="&username&"&cid="&cid&"'>"&cname&"</a>&nbsp;&nbsp;"
	rstemp2.MoveNext
  loop
  rstemp2.close
  sql="select top 1 * from favorite where username='"&username&"'"
  rstemp2.open sql,conn,3
  if rstemp2.eof then
  	showfav=0
  else
  	showfav=1
  end if
  rstemp2.close
  set rstemp2=nothing
  %>
    <td align="left">&nbsp;&nbsp;<a href="bsblog.asp?user=<%=trim(username)%>"><%=show%>的Blog</a><%if showfav=1 then%>&nbsp;&nbsp;<a href="showfavorite.asp?user=<%=trim(username)%>"><%=show%>的收藏</a><%end if%>&nbsp;&nbsp;<%=showbanner%><a href="display.asp?user=<%=username%>">关于<%=show%></a>&nbsp;&nbsp;<a href="mailto:<%=email%>">EMAIL</a></td>
  </tr>
</table>
