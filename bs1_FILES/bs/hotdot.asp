<!--#include file="../VAR.asp"-->
<!--#include file="publicmodule/function.asp"-->
<!--#include file="forum/aspFunctions.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<LINK href="../images/define.css" type=text/css rel=stylesheet>
<title>hotdot</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body leftmargin="0" topmargin="0">
  <%
   set conn=server.CreateObject("ADODB.Connection")
  'Response.Write openstring
  'Response.End 
  conn.Open OpenString
  set rs=server.CreateObject("ADODB.Recordset")
	w1=0.9
	w2=-0.06
	w3=-0.1
	w4=0.2
	'sql="select top 10  * from articles where fatherid=0 order by "&w1&"*datediff('d',ondate,"&date&")+"&w2&"*readtimes+"&w3&"*children+"&w4&"*datediff('d',lastdate,"&date&")"
	sql="select top 10 * from articles where fatherid=0 order by lastdate desc"
	'Response.Write sql
	'Response.End 
	rs.Open sql,conn,3
	if rs.EOF then
		Response.Write "<div align=center>对不起，目前没有发表的文章</div>"
	else	
  total=0
	do while not rs.EOF 
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
		 end if
		 rs1.close
		 mark1=""
		 if rs("children")<>0 then
			mark1="最后回复:" &replyauthor& "&nbsp;"& rs("lastdate")
			if len(mark1)>30 then
				mark1=left(mark1,30)&"..."
			end if
         end if
		 author=trim(rs("author"))
         if not author="" then
                 temp=split(author,"(")
                 username=temp(0)
         end if  
         if right(author,2)="()" then
			author=left(author,len(author)-2)
		 end if
		 ttitle=trim(rs("title"))
		 if len(ttitle)>10 then
			ttitle=left(ttitle,10)&"..."
		 end if
		 sql="select * from memberdes where username='"&author&"'"
		 set rs2=server.CreateObject("ADODB.Recordset")
		 rs2.open sql,conn,3
		 if not rs2.eof then
			 portrait=rs2("portrait")
		 end if
		 rs2.close
		 set rs2=nothing
		 %>
<UL class=tlst>
  <LI class=nlst>
  <H3><A style="FLOAT: left" href="viewforum.asp?id=<%=id%>&UrlTail=<%=MyUrlEncode(UrlTail)%>" target="_blank" title="作者:<%=username%>"><%=ttitle%></A>　　<SPAN class=date><%=mark1%></SPAN> 
  </H3></LI>
  <LI class=ilst style="CLEAR: both"><A href="membership/display.asp?user=<%=username%>" target="_blank"><IMG class=fil width=48 height=48 src="<%=ShowPortrait(portrait)%>" alt="<%=username%>"></A></LI>
  <LI class=clst>
  <DIV 
  id=review_1427629_short><%=GetAbstract(rs("content"),60)&"..."%> 
  <A class=pl href="viewforum.asp?id=<%=id%>&UrlTail=<%=MyUrlEncode(UrlTail)%>" target="_blank">(<%=rs("children")%>回应)</A></DIV>
  <DIV id=review_1427629_full style="DISPLAY: none"></DIV></LI></UL>
<DIV class=clear></DIV>
  <%rs.MoveNext
          loop%>
</table>
<%end if
rs.close%>
</body>
</html>

