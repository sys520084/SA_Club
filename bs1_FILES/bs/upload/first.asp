<!--#include file="..\..\var.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><!-- InstanceBegin template="/Templates/bs.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>�ϴ��ļ�</title>
<!-- InstanceEndEditable --> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<script language="javascript">
<!--
function js_callpages(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=600,height=430");
  newwin.focus();
  return false;
}
-->
</script>
<!-- InstanceEndEditable --> 
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:white;}
a:link {text-decoration: none; color: yellow }
a:visited {
	text-decoration: none;
	color: #FFFF00;
}
a:hover {
	text-decoration: underline;
	color: #FF9900;
}
a:active {text-decoration;
	color: #FF9900;
}
.unnamed1 {  font-family: "����"; font-size: 9pt; font-style: normal; line-height: 100%; font-weight: normal; color: #99FFFF}
.unnamed2 {  font-family: "����"; font-size: 9pt; line-height: 100%; font-weight: normal; color: #294C39}
input {
	font-family: "����";
	font-size: 10pt;
	color: #FFFFFF;
	background-color: #000033;
}
textarea {
	font-family: "����";
	font-size: 9pt;
	color: #FFFFFF;
	background-color: #003366;
}
-->
</style>
</head>

<body bgcolor="#043F80" link="#00FFFF" alink="#00FFFF" leftmargin="0" topmargin="0">
<p align="center"></p>
<!-- InstanceBeginEditable name="content" --><div align=center>
<%'Response.End 
  session("loadtype")=""
  if session("user")="" then
         Response.Write "<div align='center'>�Բ���������ʹ���ϴ����񣬿���������ԭ��<br>1.���ľ���ֵ�������ϴ���Ҫ����<font color=red>"&ALLOW_UP_VALUE&"</font>�ľ���ֵ����Ͽ췢������ȡ�㹻�ľ���ֵ��<br>2.�������Ѿ�����������ϴ��ļ���Ŀ�����������ԣ�<br>3.����������Ҫ������֤����<a href='../membership/relogin.asp' onclick='return js_callpage(this.href)'>���µ�½</a>��</div>"
         Response.End
  end if
  'Response.Write request("type")
 
if request("type")<>"articles" then
  set conn=server.CreateObject("ADODB.Connection")
  set rs=server.CreateObject("ADODB.Recordset")
  conn.Open OpenString
  sql="select * from files where title=''"
  rs.Open sql,conn,3
   
  do while not rs.eof
	id=rs("id")
	filename=rs("filename")
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	'Set objFSO = Server.CreateObject("Scripting.FMG1TYKZ")
	if filename<>"" then
	 if objFSO.fileExists(WebAddress()&UPFILEPATH&filename) then
	      objFSO.DeleteFile(WebAddress()&UPFILEPATH&filename)
	 end if
	end if
	sql="delete from files where id="&id
	conn.Execute sql
	rs.MoveNext
 loop

 rs.Close 
	
  user=session("user")
  sql="select count(*) as cn from files where username='"&user&"' and uploaddate=#"&date&"#"
  sql1="select sum(bytes) as sm from files where username='"&user&"'"
  'Response.Write sql
  'Response.End 
  
  rs.Open sql,conn,3
  count=0
  size=0
  if not rs.EOF then
    count=rs("cn")
    rs.Close 
    rs.Open sql1,conn,3
    if not rs.EOF then
		if not isnull(rs("sm")) then
			size=rs("sm")
		end if
	end if
  end if
	rs.Close 
	'Response.Write "ok"&size
  %>
  <p><strong>�û�<%=user%>�ϴ��ļ�����</strong></p>
  <p>��ÿ���ϴ��ļ���������<%=DAY_FILES%>�������������<font color=red><%=MAX_FULLSPACE/1000%> K</font>�������ڻ���<font color=red><%=int((MAX_FULLSPACE-size)*100/1024)/100%> K</font>�ռ䣩</p>
  <%
     pageNo=request("page")
     PgSz=20

    sql="select * from files where username='"&user&"' order by id desc"
     rs.Open sql,conn,3	 
     rs.PageSize=cint(PgSz)
     all_page=rs.PageCount
     total=rs.RecordCount 
     if total<>0 then
         if pageNo<>"" then
            rs.AbsolutePage =pageNo
            RcoNo=1
         else
            rs.AbsolutePage =1
            rs.movefirst
            pageNo=1 
         end if%>
<div align="center"> 
  <%if pageNo>1 then%>
  <a href="first.asp?page=<%=pageNo-1%>">��һҳ</a>&nbsp; 
  <%end if%>
  <%if pageNo-all_page<0 or pageNo="" then%>
  <a href="first.asp?page=<%=pageNo+1%>">��һҳ</a> 
  <%end if%>
</div>

  <table width="100%" border="1" cellpadding="1" cellspacing="1" bordercolor="#6666FF">
    <tr align="center" bgcolor="#000033"> 
      <td width="13%">�ļ���</td>
      <td width="25%">�ļ�����</td>
      <td width="10%">�ϴ�����</td>
      <td width="8%">���ش���</td>
      <td width="5%">��СK</td>
      <td width="10%">����</td>
    </tr>
    <%for iPage=1 to rs.Pagesize
	          RecNo=(pageNo-1)*PgSz+iPage%>
    <tr bgcolor="#6666FF"> 
      <td align="center"><a href="download.asp?id=<%=rs("id")%>"><%=rs("filename")%></a></td>
      <td><font color="#000000"><a href="detail.asp?id=<%=rs("id")%>&url=first.asp&thispage=<%=pageNo%>"><%=rs("title")%></a></font></td>
      <td align="center"><font color="#000000"><%=rs("uploaddate")%></font></td>
      <td align="center"><font color="#000000"><%=rs("times")%></font></td>
      <td align="center"><font color="#000000"><%=int(rs("bytes")*100/1024)/100%></font></td>
      <td align="center"><font color="#000000"><a href="savedoc.asp?id=<%=rs("id")%>" onclick="return js_callpages(this.href)">�༭</a>&nbsp;
      <a href="delete.asp?id=<%=rs("id")%>" onclick="if(confirm('ȷʵɾ������ļ�ô��'))return js_callpages(this.href);">ɾ��</a> </font></td>
      <%        rs.MoveNext
               if RecNo=total or rs.eof then
                 exit for
               end if%>
    </tr>
    <%next%>
  </table>
  <%   
     if pageNo<>"" then
        response.write "&nbsp;&nbsp;<font align=center> ��"&pageNo&"ҳ/ ��"&all_page&"ҳ</font>"
     else
        response.write "&nbsp;&nbsp;<font align=center> ��1ҳ / ��"&all_page&"ҳ</font>"   
     end if
     rs.close%>
  &nbsp;&nbsp;�鿴�� 
  <%for action=1 to all_page%>
  <a href="first.asp?page=<%=action%>">�� <%=action%> ��</a> 
  <%next%>
 <%else%>
  <div align=center>û���ҵ��κ���Ϣ</div>

  <%end if
else
	session("loadtype")="articles"
end if 'end if type<>"articles"
  if count>=DAY_FILES then
	Response.Write "<div align=center>���Ѿ�������ÿ���ϴ�������ļ������������ϴ��ˣ�</div>"
  else 
	if size>=MAX_FULLSPACE then
		Response.Write "<div align=center>���Ѿ�����������ӵ�е�����������������ϴ��ˣ�</div>"
	else
  %>
  
  <p><b>��һ�����ϴ��ļ�</b></p>
  <form method="post" name="form1" action="savedoc.asp" enctype="multipart/form-data">
    <p>��ѡ��Ҫ�ϴ����ļ��� 
      <input type="file" name="doc">
      ���ļ��������������ģ��� 
      <input type="submit" name="confirm" value="ȷ��">
    </p>
    </form>
  <%end if
  end if
%>
</div><!-- InstanceEndEditable -->
<p align="center"><font color="#CCFF66">���Ǿ��ֲ�����Ȩ���� <font face="Arial, Helvetica, sans-serif"><br>
  Copyright 2002-2003 Clustering Intelligence Club</font></font> </p>
</body>
<!-- InstanceEnd --></html>