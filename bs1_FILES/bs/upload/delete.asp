<!--#include file="..\..\var.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
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
-->
</style>
</head>

<body bgcolor="#043F80" link="#00FFFF" alink="#00FFFF" leftmargin="0" topmargin="0">
<div align=center>
<%
if session("user")="" then
         Response.Write "<div align='center'>�Բ���������ʹ���ϴ�������<a href='../membership/relogin.asp' onclick='return js_callpage(this.href)'>���µ�½</a>��</div>"
         Response.End
  end if
  user=session("user")
  id=request("id")
  sql="select * from files where id="&id
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  set rs=server.CreateObject("ADODB.Recordset")
  sql="select * from files where id="&id
  
  rs.Open sql,conn,3
  if rs.EOF then
	Response.Write "<div align='center'>�Բ������ݲ�����������û�ҵ�ָ���ļ�¼</div>"
	Response.End
  end if
  filename=rs("filename")
  rs.Close
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
  'Set objFSO = Server.CreateObject("Scripting.FMG1TYKZ")
  if filename<>"" then
   if objFSO.fileExists(WebAddress()&UPFILEPATH&filename) then
        objFSO.DeleteFile(WebAddress()&UPFILEPATH&filename)
   else
        response.write "δ�ҵ�"&filename
   end if
  end if
  sql="delete from files where id="&id
  'Response.Write sql
  'Response.End
  conn.Execute sql	
%>
	ɾ���ļ��ɹ���</div>
</body>
</html>
