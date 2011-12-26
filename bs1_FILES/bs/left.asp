<!--#include file="../VAR.asp"-->
<!--#include file="forum/aspFunctions.asp"-->
<!--#include file="publicmodule/function.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>blog</title>
<meta http-equiv="refresh" content="600">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/define.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	background-image: url(images/leftbg_03.gif);
}
-->
</style></head>
<body>
<%
user=session("user")
set conn=server.CreateObject("ADODB.Connection")
conn.Open OpenString
'set rs=server.CreateObject("ADODB.Recordset")
'sql="select * from forum_state where fatherid<>0"
'rs.Open sql,conn,3
isadmin=IsSuperAdmin(Conn,Session("user"))
%>
<base target="mainFrame">
<%		msgCount=0
	fullMsgCount=0
	if session("user")<>"" then
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
			<span class="blogname">
		  <%if user<>"" then%><%=user%> 的BLOG<%else%>过客参观<%end if%></span>
			<%if msgCount>0 then%>
			<img src="img/envolop.gif" alt="您有新的短消息">
			<%end if%>	  </td>
	</tr>
</table>
<hr size="1" />

<table width="90%" align="center">
	<tr>
		<td>
			&nbsp;&nbsp;<img src='images/dot_father_012.gif' /> <b>
			<font color="#333333">文章操作</font></b>
		</td>
	</tr>
</table>

<table width="90%" align="center">
<tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /></font> 
      <a href="forum/viewtitle.asp?order=<%=server.URLEncode("lastdate desc")%>" target="mainFrame">全部文章</a></td>
  </tr>
<%if session("user")="" then%>
 <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="membership/relogin.asp?UrlTail=<%=MyUrlEncode("../forum.asp")%>" target="_parent">登陆</a></font></td>
  </tr>
<%end if%>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="membership/bsblog.asp?user=<%=user%>" target="mainFrame">我的Blog</a></font></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="membership/oldbsblog.asp?user=<%=user%>&reply=<%=user%>" target="mainFrame">参与的讨论</a></font></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="forum/new.asp" target="mainFrame">写文章</a></font></td>
  </tr>
  <tr>
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /></font> <a href="forum/show-all-elites.asp?score=1&order=<%=server.URLEncode("id desc")%>" target="mainFrame">全部精华文章</a></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /></font> 
      <a href="forum/viewfiles.asp" target="mainFrame">全部上传文件</a></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="forum/finder.asp" target="mainFrame">查找文章</a></font></td>
  </tr>
</table>

<table width="90%" align="center">
  <tr>
    <td>&nbsp;&nbsp;<img src='images/dot_father_012.gif' /> <b><a href="../bs/forum.asp?user=<%=user%>" target="_blank"><font color="#333333">用户服务</font></a></b> 
    </td>
  </tr>
  <tr>
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> <a href="feed.asp">RSS订阅</a> </font></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="message/msgman.asp" target="mainFrame">消息服务</a></font></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="upload/first.asp" target="mainFrame">上传文件</a></font></td>
  </tr>
  <%if isadmin then%>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="http://www.swarmagents.cn/thesis/admin/index.asp" target="mainFrame">管理论文</a> </font></td>
  </tr>
  <tr>
     <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
	 <a href="http://www.swarmagents.com/bottomshow/manage.asp" target="mainFrame">管理Bottom</a> </font></td>
  <tr><td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="http://www.swarmagents.com/thesis/manage.asp" target="mainFrame">管理Thesis</a> </font></td>
  </tr>
  <%end if%>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="membership/changepass.htm" target="mainFrame">更改密码</a> 
      </font></td>
  </tr>
  <!--
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="membership/memeber.asp?type=%22change%22" target="mainFrame">更改详细资料</a> 
      </font></td>
  </tr>-->
  <tr>
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> <a href="membership/memeber.asp?type=%22change%22" target="mainFrame">更改基本信息</a></font></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="membership/persenal.asp" target="mainFrame">更改个人说明</a></font></td>
  </tr>
</table>

<table width="90%" align="center">
  <tr> 
    <td> &nbsp;&nbsp;<img src='images/dot_father_012.gif' /> <b><a href="../bs/index.asp" target="_parent"><font color="#333333">离开</font></a></b> 
    </td>
  </tr>
</table>
		</td>
	</tr>
</table>
</body>
</html>

