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
<title><%=username%>�ľ���Ŀ¼</title>
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
}</STYLE>
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
set rs=server.CreateObject("ADODB.Recordset")
set conn=server.CreateObject("ADODB.Connection")
Conn.open OpenString

function ShowSubTree(byval id)
	sql="select * from usercategory where fatherid="&id
	   set rs1=conn.Execute(sql)
	do while not rs1.eof
		Tid=rs1("id")
		Tcatalog_name=trim(rs1("category"))
		Tfather_id=trim(rs1("fatherid"))%>
</p>
	<ul>
    <li><a href=# onclick="javascript: givevalue('<%=Tid%>')"><%=Tcatalog_name%></a> 
      <%ShowSubTree Tid%>
    </li>
  </ul>
  <%rs1.movenext
	loop
	rs1.close
	set rs1=nothing
end function
sql="select * from articles where id="&request("articleid")
rs.open sql,conn,3
if not rs.eof then
	title=rs("title")
end if
rs.close

if request("categoryid")="" then
sql="select * from article_category,usercategory where article_id="&request("articleid") &" and usercategory.id=article_category.category_id"
'response.write sql
'response.end
rs.open sql,conn,3
if not rs.eof then
%>
<div align=left>����<font color=red><%=title%></font>�Ѿ�����Ŀ¼��</div>
<ul>
<% do while not rs.eof%>
<li><div><%=rs("category")%><a href="delElite.asp?relationid=<%=rs(0)%>&articleid=<%=request("articleid")%>"><img border=0 src="../images/delete.gif" alt="�����Ŀ¼��ɾ����ƪ����"></a></div></li>
<% rs.movenext
loop%>
</ul>
<hr width=90% align=center>
<%end if
rs.close
end if%>
<div align=left>������<font color=red><%=title%></font>����Ŀ¼������������Ŀ¼���ƣ�</div>

<%

sql="Select * From usercategory where fatherid=-1 and username='"&username&"'"
  'Response.Write sql 
  rs.Open sql,Conn,3
  if rs.eof then%>
  <p align="center">����ǰû���κξ���Ŀ¼</p>
<%end if
   do while not rs.eof 
  		Tid=trim(rs("id"))
		Tcatalog_name=trim(rs("category"))
		Tfather_id=trim(rs("fatherid"))%>
  <ul>
    <li><a href=# onclick="javascript: givevalue('<%=Tid%>')"><%=Tcatalog_name%></a> 
      <%ShowSubTree Tid%>
    </li>
  </ul>
  <p> 
    <% rs.movenext
	  loop
  %>
  </p>
  <form action="addware.asp" method="post" name="mainfrm" id="mainfrm">
  <input name="categoryid" type="hidden" value="">
  <input name="ocategoryid" type="hidden" value="<%=request("categoryid")%>">
  <input name="articleid" type="hidden" value="<%=request("articleid")%>">
</form>
<hr width=90% align=center>
<div align=left>�½�����Ŀ¼������<font color=red><%=title%></font>�����½���Ŀ¼</div>
<form name="newForm" method="post" action="addware.asp">
<%
rs.close
sql="select * from usercategory where username='"&username&"'"
rs.open sql,conn,3
%>
  <p align="center">�� 
    <select name="fatherid" id="fatherid">
      <option value="-1">��Ŀ¼</option>
      <%do while not rs.eof%>
      <option value="<%=rs("id")%>"><%=rs("category")%></option>
      <%rs.movenext
loop
rs.close
set rs=nothing%>
    </select>
    Ŀ¼�£��½�һ��Ŀ¼: 
    <input type="text" name="categoryname">
    <input type="hidden" name="articleid" value="<%=request("articleid")%>">
   <input name="ocategoryid" type="hidden" value="<%=request("categoryid")%>">
	<input type="hidden" name=
  </p>
  <p align="center">�½�Ŀ¼���ܣ�</p>
  <p align="center">
    <textarea name="descript" cols="50" rows="5" id="descript"></textarea>
  </p>
  <p align="center">
    <input type="submit" name="Submit" value="Submit">
  </p>
  </form>
<p align="center">&nbsp;</p>
</body>
</html>
