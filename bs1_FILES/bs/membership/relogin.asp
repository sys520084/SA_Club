<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<!--#include file="../forum/aspFunctions.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>�û���½</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>DIV {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
}
TD {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
}
A {
	COLOR: #0099FF; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: none
}
A:hover {
	COLOR: #ff0000; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: underline
}
</STYLE>
</head>
<body>
<%'Response.Write request("login")
  if request("login")="ȷ��" then
     username=SqlTran(request("name"))
     if username="" then
        Response.Write "<div align=center>�Բ�����û�������û�����<a href='javascript:history.go(-1)'>��������</a></div>"
        Response.End
     end if
     set conn=server.CreateObject("ADODB.Connection")
     conn.Open OpenString
     sql="select passwd from users where username='"&username&"'"
     set rs=conn.Execute(sql)
     if rs.eof then
        Response.Write "<div align=center>�Բ�����������û�������ȷ.<a href='javascript:history.go(-1)'>��������</a></div>"
        Response.End
     end if
     if trim(rs("passwd"))<>trim(request("pass")) then
        Response.Write "<div align=center>�Բ�������������벻��ȷ.<a href='javascript:history.go(-1)'>��������</a></div>"
        Response.End
     end if
     session("user")=username
     rs.close
     '��ȡȨ��
     sql="select * from power where username='"&username&"' and (power=0 or power=2)" 
     rs.Open sql,conn
	 if not rs.eof then
	 	session("database")="thesis"
	 end if
	 UrlTail=MyUrlDecode(request("UrlTail"))
     %>
<div align="center">���Ѿ���ȷͨ���������֤����<a href="<%=UrlTail%>">��������ʱ���ҳ��</a>��</div>
<%	'UrlTail=MyUrlDecode(request("UrlTail"))%>
<!--<script language="javascript">
  //window.location.replace("<%=UrlTail%>");
  //url=window.opener.location;
  //window.opener.location.reload();
  //window.close();
</script>-->
<%end if%>
<%if Request.ServerVariables("REQUEST_METHOD")<>"POST" then%>     
<div align="center">
  <p>�����ڳ�ʱ��û��ʹ�ñ�ϵͳ�Ѿ��˳��������½��������֤��<a href="fastaddnew.asp?UrlTail=<%=request("UrlTail")%>" target="_blank">����ע�����û�</a>��</p>
  <form name="frm1" action="relogin.asp" method="POST">
    <p>�û���½��</p>
    <p>�û����� 
      <input type="text" name="name">
    </p>
    <p>���룺 &nbsp; 
      <input type="password" name="pass">
    </p>
    <p> 
	  <input type="hidden" name="UrlTail" value="<%=request("UrlTail")%>">

      <input type="submit" name="login" value="ȷ��">
      <input type="reset" name="Submit2" value="��д">
    </p>
  </form>
  <a href="fastaddnew.asp?UrlTail=<%=request("UrlTail")%>"><font color="#0000FF">����ע�����û���</font></a></div>
<%end if%>
</body>
</html>
