<div align="center"><p>&nbsp;</p>
	<%
	set rstemp3=server.CreateObject("ADODB.Recordset")
	sql="select top 20 * from favorite where username='"&request("user")&"' order by id desc"
	'response.write sql
	'response.end 
	'response.write conn
	
	rstemp3.open sql,conn,3
	if not rstemp3.eof then%>
	 <p  class="bs_blog_title"> ’≤ÿ±Í«©</p>
<%  end if
	'response.end
	dim tags(50)
	 
	count=1
	do while not rstemp3.eof
		tagrs=rstemp3("tag")&" "
		'response.write tagrs
		if tagrs<>"" then
			tagdims=split(tagrs," ")
			for i=0 to ubound(tagdims)-1
			'response.write tagdims(i)
				found=false
				for j=0 to count-1
					if tags(j)=tagdims(i) then
						found=true
						exit for
					end if
				next
				if not found then
					'redim Preserve tags(count+1)
					if count<=50 then
						tags(count)=tagdims(i)
						count=count+1
					end if
				end if
				'response.write ubound(tags)
			next
		end if
		rstemp3.movenext
	loop
	rstemp3.close
	set rstemp3=nothing
	for i=0 to count-1
	%>
	<div align="center"><p><a href="showfavorite.asp?tags=<%=tags(i)%>&user=<%=request("user")%>"><%=tags(i)%></a></p></div>
	<%
	next
	%>
</div>