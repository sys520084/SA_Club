<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<%username=session("user")
if username="" or isnull(username) then
	response.write "�㻹û�е�½�����ܴ򿪱�ҳ��"
	response.end
end if
%>
<title><%=username%>���ղ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
DIV,p {FONT-FAMILY: "����"; FONT-SIZE: 9pt; color:#000000;}
TD {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
}
A {COLOR: #0099FF; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #FF0000; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
input,textarea,SELECT {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
	background-color: #FFFFFF;
}.style1 {color: #FF0000}
.style2 {color: #000000}
</STYLE>
<script>
function givevalue(val)
{
	mainfrm.categoryid.value=val;
	mainfrm.submit();
}
</script>
</head>

<body>
<p>
  <%
id=request("id")
if id="" then
	response.write "��û��ָ��Ҫ�ղص����£����ܲ�����"
	response.End
end if  
if session("user")="" then
		 UrlTail=MyUrlEncode("../elite/favorite.asp?id=" & request("id"))
         Response.Write "<div align='center'>�Բ��������ܷ������£���<a href='../membership/relogin.asp?UrlTail=" & UrlTail & "'>���µ�½</a>�����ߣ�</div>"
         Response.End
 end if
set rs=server.CreateObject("ADODB.Recordset")
set conn=server.CreateObject("ADODB.Connection")
Conn.open OpenString
if request("submit")<>"ȷ��" then
	sql="select * from favorite where article_id="&Id
	'response.write sql
	'response.end 
	rs.open sql,conn,3
	if not rs.eof then
		response.write "���Ѿ��ղ���ƪ�����ˣ������ظ��ղأ�"
		response.End
	end if
	rs.close
	sql="select * from articles where id="&id
	'response.write sql
	'response.end 
	rs.open sql,conn,3
	'response.end
	if rs.eof then
		response.write "û���ҵ���Ҫ�ղص����£���ȷ�ϴ������Ƿ��Ѿ���ɾ����"
		response.End
	end if
	
	title=rs("title")
	author=rs("author")
	rs.close
	set rs=nothing
%>
<form name="form1" method="post" action="">
  <p>��Ҫ�ղ�<span class="style1"><%=author%><span class="style2">��</span><%=title%></span></p>
  <p>��Ϊ��Ҫ�ղص���ƪ������ӱ�ǩ˵������ǩ�Կո����:</p>
  <p>
    <input name="id" type="hidden" id="id">
  </p>
  <p>
    <input name="tags" type="text" id="tags" size="50">
    <input type="submit" name="Submit" value="ȷ��">
    </p>
</form>
<%
else
	tags=MyhtmlEncode(request("tags"))
	username=session("user")
	ondate=date()&" "&time()
	sql="insert into favorite (username,article_id,tag,ondate) values ('"&username&"',"&id&",'"&tags&"',#"&ondate&"#)"
	'response.write sql
	conn.execute(sql)
	conn.close
	set conn=nothing
	%>
���Ѿ��ɹ��ղ�1ƪ���£�
	
	<%
end if
%>
<p><p align="center">&nbsp;</p>
</body>
</html>
