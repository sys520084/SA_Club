<!-- #include file="../../var.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:����;color:black;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0000CC}
a:visited {text-decoration: none;; color: #660066}
a:hover {text-decoration: underline; color: #FF0000}
a:active {text-decoration;background-color: #CCCCCC; color: #006666}
-->
</style>
<title>�����°�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<%if not session("admin") then
     Response.Write "<div algin='center'>�Բ��������ܴ򿪱�ҳ��<a href='jingwen.asp'>�����µ�½</a>"
     response.end
  end if
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  set rs=server.CreateObject("ADODB.Recordset")
  forum=request("forum")
  fatherid=cint(request("father"))
  title=cint(request("title"))
  if forum="" then
     Response.Write "<div align=center>Ҫ�ӵ���̳���Ʋ���Ϊ��</div>"
     Response.End
  end if 
  rs.Open "select * from forum_state where catagory='"&forum&"'",conn
  if not rs.EOf then
     Response.Write "<div align=center>Ҫ�ӵ���̳������������̳����</div>"
     Response.End
  end if   
  rs.Close
  if fatherid<>0 then
	rs.Open "select * from forum_state where id="&fatherid
	if not rs.EOF then
		if cint(rs("fatherid"))<>0 then
			 Response.Write "�����������̳���������̳��"
			 Response.End
		end if
	end if
	rs.Close 
  end if
  rs.Open "forum_state",conn,1,3
  rs.AddNew
  rs("catagory")=trim(forum)
  rs("fatherid")=fatherid
  rs("child")=title
  rs.Update
  rs.Close
  %>   
<body bgcolor="#FFFFFF">
<div align="center">
  <p>�����°�ɹ���</p>
</div>
</body>
</html>
