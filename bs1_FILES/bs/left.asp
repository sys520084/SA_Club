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
		  <%if user<>"" then%><%=user%> ��BLOG<%else%>���Ͳι�<%end if%></span>
			<%if msgCount>0 then%>
			<img src="img/envolop.gif" alt="�����µĶ���Ϣ">
			<%end if%>	  </td>
	</tr>
</table>
<hr size="1" />

<table width="90%" align="center">
	<tr>
		<td>
			&nbsp;&nbsp;<img src='images/dot_father_012.gif' /> <b>
			<font color="#333333">���²���</font></b>
		</td>
	</tr>
</table>

<table width="90%" align="center">
<tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /></font> 
      <a href="forum/viewtitle.asp?order=<%=server.URLEncode("lastdate desc")%>" target="mainFrame">ȫ������</a></td>
  </tr>
<%if session("user")="" then%>
 <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="membership/relogin.asp?UrlTail=<%=MyUrlEncode("../forum.asp")%>" target="_parent">��½</a></font></td>
  </tr>
<%end if%>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="membership/bsblog.asp?user=<%=user%>" target="mainFrame">�ҵ�Blog</a></font></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="membership/oldbsblog.asp?user=<%=user%>&reply=<%=user%>" target="mainFrame">���������</a></font></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="forum/new.asp" target="mainFrame">д����</a></font></td>
  </tr>
  <tr>
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /></font> <a href="forum/show-all-elites.asp?score=1&order=<%=server.URLEncode("id desc")%>" target="mainFrame">ȫ����������</a></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /></font> 
      <a href="forum/viewfiles.asp" target="mainFrame">ȫ���ϴ��ļ�</a></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="forum/finder.asp" target="mainFrame">��������</a></font></td>
  </tr>
</table>

<table width="90%" align="center">
  <tr>
    <td>&nbsp;&nbsp;<img src='images/dot_father_012.gif' /> <b><a href="../bs/forum.asp?user=<%=user%>" target="_blank"><font color="#333333">�û�����</font></a></b> 
    </td>
  </tr>
  <tr>
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> <a href="feed.asp">RSS����</a> </font></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="message/msgman.asp" target="mainFrame">��Ϣ����</a></font></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="upload/first.asp" target="mainFrame">�ϴ��ļ�</a></font></td>
  </tr>
  <%if isadmin then%>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="http://www.swarmagents.cn/thesis/admin/index.asp" target="mainFrame">��������</a> </font></td>
  </tr>
  <tr>
     <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
	 <a href="http://www.swarmagents.com/bottomshow/manage.asp" target="mainFrame">����Bottom</a> </font></td>
  <tr><td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="http://www.swarmagents.com/thesis/manage.asp" target="mainFrame">����Thesis</a> </font></td>
  </tr>
  <%end if%>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="membership/changepass.htm" target="mainFrame">��������</a> 
      </font></td>
  </tr>
  <!--
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="membership/memeber.asp?type=%22change%22" target="mainFrame">������ϸ����</a> 
      </font></td>
  </tr>-->
  <tr>
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> <a href="membership/memeber.asp?type=%22change%22" target="mainFrame">���Ļ�����Ϣ</a></font></td>
  </tr>
  <tr> 
    <td><font color="#333333">&nbsp;&nbsp;&nbsp;<img src="images/dot_child_01.gif" /> 
      <a href="membership/persenal.asp" target="mainFrame">���ĸ���˵��</a></font></td>
  </tr>
</table>

<table width="90%" align="center">
  <tr> 
    <td> &nbsp;&nbsp;<img src='images/dot_father_012.gif' /> <b><a href="../bs/index.asp" target="_parent"><font color="#333333">�뿪</font></a></b> 
    </td>
  </tr>
</table>
		</td>
	</tr>
</table>
</body>
</html>

