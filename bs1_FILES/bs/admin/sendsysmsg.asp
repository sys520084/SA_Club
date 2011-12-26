<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>refresh</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:宋体;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
-->
</style></head>

<body bgcolor="#333366" text="#FFFFFF" leftmargin="0" topmargin="0">
<%'Response.End 
if not session("admin") then
     Response.Write "<div algin='center'>对不起，您不能打开本页。<a href='jingwen.asp'>请重新登陆</a>"
     response.end
  end if
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  if Request.ServerVariables("REQUEST_METHOD")="POST" then
		reciever=request("reciever")
		if request("reciever")="" then
			reciever="all"
		end if
		content=request("content")
		title=request("title")
		title=replace(title,"'","''")
		content=replace(content,"'","''")
		sql="insert into message (sender,reciever,senddate,sendtime,type,title,content) values"&_
				"('system','"&reciever&"','"&date&"','"&time&"',0,'"&title&"','"&content&"')"
		conn.Execute sql
  		%>
  		<div align=center>发送消息成功</div>
  <%else
  sql="select * from message where reciever='system' order by id desc"
  set rs=server.CreateObject("ADODB.Recordset")
  rs.Open sql,conn,3 %>
<div align="center">
  <table width="90%" border="1" cellpadding="1" cellspacing="1" bordercolor="#0066FF" bordercolorlight="#666666" bordercolordark="#CCCCCC">
    <tr bgcolor="#000066" align="center"> 
      <td width="35%"><font color="#CCCCFF">题目</font></td>
      <td width="15%"><font color="#CCCCFF">发送者</font></td>
      <td width="20%"><font color="#CCCCFF">发送时间</font></td>
      <td width="15%"><font color="#CCCCFF">操作</font></td>
    </tr>
    <%do while not rs.eof%>
    <tr bgcolor="#6666FF" align="center"> 
      <td width="50%"><a href="viewmsg.asp?id=<%=rs("id")%>" onclick="return js_callpage(this.href)"><%=trim(rs("title"))%></a> 
        <%if trim(rs("content"))="" or isnull(rs("content")) then%>
        <font color="#000000">（内空） </font>
		<%end if%>
        &nbsp; 
        <%if rs("status")=0 then%>
        <font color="red">新！</font> 
        <%end if%>
      </td>
      <td width="15%"><a href="../membership/display.asp?name=<%=trim(rs("sender"))%>"  onclick="return js_callpage(this.href)"><%=trim(rs("sender"))%></a></td>
      <td width="20%"><%=trim(rs("senddate"))&"&nbsp;"&trim(rs("sendtime"))%></td>
      <td width="5%"> 
        <a href="../message/sendmsg.asp?name=<%=trim(rs("sender"))%>&id=<%=rs("id")%>&user=<%="system"%>"><img src="../img/re.gif" alt="回复发送者" width="14" height="17" border="0"></a>&nbsp;<a href="../message/delmsg.asp?id=<%=rs("id")%>&user=system" onclick="return delmsg(this.href);"><img src="../img/del.gif" alt="删除" width="15" height="15" border="0"></a>
        <%     rs.MoveNext
      loop%>
    </tr>
  </table>
  <%end if%>
</body>
</html>
