<!--#include file="../VAR.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>refresh</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<meta http-equiv="refresh" content="1200">
<link href="css/define.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	background-image: url(images/leftbg_03.gif);
}
-->
</style>
</head>



<body>
<%		msgCount=0
	fullMsgCount=0
	user=session("user")
	if session("user")<>"" then
		set conn=server.CreateObject("ADODB.Connection")
		conn.Open OpenString
	
		sql="select count(*) from message where (reciever='"&session("user")&"' or (type=0 and reciever='all'))  and status=0"
		'Response.Write sql
		'Response.End 
		set rs=conn.Execute(sql)
		msgCount=0
		if not rs.eof then
			msgCount=rs(0)
		end if
		rs.close
		sql="select count(*) from message where reciever='"&session("user")&"' or (type=0 and reciever='all') "
		set rs=conn.Execute(sql)
		if not rs.eof then
			fullMsgCount=rs(0)
		end if
		rs.close
		sql="select * from online where username='"&session("user")&"'"
		rs.open sql,conn,1,3
		if not rs.eof then
			time1=rs("onlinetime")
			time2=now()
			sec=datediff("s",time1,time2)
			if sec>300 then
				addsec=0
			else
				addsec=sec
			end if
			rs("onlinetime")=time2
			rs("sumtime")=rs("sumtime")+addsec/3600
			rs.update
		end if
		rs.close
	end if
%>
<table cellspacing="0" cellpadding="0" border="0">
	<tr>
		<td align="center">
			<img src="images/logo_black_01.gif">
			<span class="blogname"><%if user<>"" then%><%=user%> 的BLOG<%else%>过客参观<%end if%></span>
			<%if msgCount>0 then%>
			<img src="img/envolop.gif" alt="您有新的短消息">
			<%end if%>
		</td>
	</tr>
</table>

</body>
</html>
