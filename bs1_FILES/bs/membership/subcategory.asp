<div align="center"><p>&nbsp;</p>
	<%
	sql="select * from usercategory where fatherid="&request("cid")
	rstemp3.open sql,conn,3
	if not rstemp3.eof then%>
	 <p  class="bs_blog_title">��Ŀ¼</p>
<%  end if	
	do while not rstemp3.eof
	%>
	<div align="center"><p><a href="showelite.asp?cid=<%=rstemp3("id")%>&user=<%=request("user")%>"><%=rstemp3("category")%></a></p></div>
	<%
		rstemp3.movenext
	loop

	if session("user")=request("user") then%>
                <form action="../elite/addware.asp" method="post" name="frmAdd" id="frmAdd" target="_blank">
				�ڵ�ǰĿ¼�£�<br>
                  <input type="text" name="categoryname">
				  <input type="hidden" name="user" value=<%=request("user")%>>
				  <input type="hidden" name="fatherid" value=<%=request("cid")%>>
				  
                  <input type="submit" name="Submit" value="�½���Ŀ¼">
                </form>
				<a href="../elite/editcategory.asp?categoryid=<%=request("cid")%>" target="_blank"><img src="../images/edit.gif" width="20" height="20" alt="�޸ĵ�ǰĿ¼" border=0></a>
				<a href="../elite/delcategory.asp?categoryid=<%=request("cid")%>" onclick="javascript:return confirm('ȷʵҪɾ�����Ŀ¼��')"  target="_blank"><img src="../images/delete.gif" width="20" height="20" alt="ɾ����ǰĿ¼" border=0></a>
<%end if%>
</div>