<div align="center"><p>&nbsp;</p>
	<%
	sql="select * from usercategory where fatherid="&request("cid")
	rstemp3.open sql,conn,3
	if not rstemp3.eof then%>
	 <p  class="bs_blog_title">子目录</p>
<%  end if	
	do while not rstemp3.eof
	%>
	<div align="center"><p><a href="showelite.asp?cid=<%=rstemp3("id")%>&user=<%=request("user")%>"><%=rstemp3("category")%></a></p></div>
	<%
		rstemp3.movenext
	loop

	if session("user")=request("user") then%>
                <form action="../elite/addware.asp" method="post" name="frmAdd" id="frmAdd" target="_blank">
				在当前目录下：<br>
                  <input type="text" name="categoryname">
				  <input type="hidden" name="user" value=<%=request("user")%>>
				  <input type="hidden" name="fatherid" value=<%=request("cid")%>>
				  
                  <input type="submit" name="Submit" value="新建子目录">
                </form>
				<a href="../elite/editcategory.asp?categoryid=<%=request("cid")%>" target="_blank"><img src="../images/edit.gif" width="20" height="20" alt="修改当前目录" border=0></a>
				<a href="../elite/delcategory.asp?categoryid=<%=request("cid")%>" onclick="javascript:return confirm('确实要删除这个目录吗？')"  target="_blank"><img src="../images/delete.gif" width="20" height="20" alt="删除当前目录" border=0></a>
<%end if%>
</div>