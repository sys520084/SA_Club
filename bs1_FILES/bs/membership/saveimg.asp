<!--#include file="../../var.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<!--#include file="../publicmodule/fileoperator.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>����ͼƬ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>DIV {
	FONT-FAMILY: "����"; FONT-SIZE: 9pt
}
TD {
	FONT-FAMILY: "����"; FONT-SIZE: 9pt
}
A {
	COLOR: #003366; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: none
}
A:hover {
	COLOR: #ff0000; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: underline
}
</STYLE>
</head>

<body bgcolor="#FFFFFF">
<div align="center">
<%myuser=session("user")
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open OpenString
	set rs=server.CreateObject("ADODB.Recordset")
	'response.write WebAddress() & "userface\"
	'response.end
	tb=UploadFile(WebAddress() & "\bs\userface\",server.URLEncode(myuser),1024*1024,MAX_FULLSPACE,"jpg,gif")
	response.write tb
	if tb<>"" then
		sql="update memberdes set portrait='"&tb&"' where username='"&myuser&"'"
		'response.write sql
		'response.end
		conn.execute sql
	end if
	%>
  <p>�ϴ�ͷ��ɹ���</p>
  <p>�û���<%=user%></p>
  <p>��Ƭ��<img src="<%=ShowPortrait(tb)%>"><br>
  <p></p>
</div>
</body>
</html>
