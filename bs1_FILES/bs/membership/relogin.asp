<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<!--#include file="../forum/aspFunctions.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>用户登陆</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>DIV {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
}
TD {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
}
A {
	COLOR: #0099FF; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none
}
A:hover {
	COLOR: #ff0000; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: underline
}
</STYLE>
</head>
<body>
<%'Response.Write request("login")
  if request("login")="确定" then
     username=SqlTran(request("name"))
     if username="" then
        Response.Write "<div align=center>对不起，您没有输入用户名！<a href='javascript:history.go(-1)'>请重来！</a></div>"
        Response.End
     end if
     set conn=server.CreateObject("ADODB.Connection")
     conn.Open OpenString
     sql="select passwd from users where username='"&username&"'"
     set rs=conn.Execute(sql)
     if rs.eof then
        Response.Write "<div align=center>对不起，您输入的用户名不正确.<a href='javascript:history.go(-1)'>请重来！</a></div>"
        Response.End
     end if
     if trim(rs("passwd"))<>trim(request("pass")) then
        Response.Write "<div align=center>对不起，您输入的密码不正确.<a href='javascript:history.go(-1)'>请重来！</a></div>"
        Response.End
     end if
     session("user")=username
     rs.close
     '读取权限
     sql="select * from power where username='"&username&"' and (power=0 or power=2)" 
     rs.Open sql,conn
	 if not rs.eof then
	 	session("database")="thesis"
	 end if
	 UrlTail=MyUrlDecode(request("UrlTail"))
     %>
<div align="center">您已经正确通过了身份验证，请<a href="<%=UrlTail%>">返回你来时候的页面</a>。</div>
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
  <p>您由于长时间没有使用本系统已经退出，请重新进行身份验证或<a href="fastaddnew.asp?UrlTail=<%=request("UrlTail")%>" target="_blank">快速注册新用户</a>：</p>
  <form name="frm1" action="relogin.asp" method="POST">
    <p>用户登陆：</p>
    <p>用户名： 
      <input type="text" name="name">
    </p>
    <p>密码： &nbsp; 
      <input type="password" name="pass">
    </p>
    <p> 
	  <input type="hidden" name="UrlTail" value="<%=request("UrlTail")%>">

      <input type="submit" name="login" value="确定">
      <input type="reset" name="Submit2" value="重写">
    </p>
  </form>
  <a href="fastaddnew.asp?UrlTail=<%=request("UrlTail")%>"><font color="#0000FF">快速注册新用户！</font></a></div>
<%end if%>
</body>
</html>
