<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <%
  username=request("user")
	show=trim(username)
  if username=session("user") then
  	show="��"
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
  %>
    <td align="center">&nbsp;&nbsp;<a href="bsblog.asp?user=<%=trim(username)%>"><%=show%>��Blog</a>&nbsp;&nbsp;<%=showbanner%><a href="display.asp?user=<%=username%>">����<%=show%></a></td>
  </tr>
</table>
