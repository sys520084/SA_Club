<!--#include file="../VAR.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>ȡ������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:����;color:black;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0000CC}
a:visited {text-decoration: none;; color: #660066}
a:hover {text-decoration: underline; color: #FF0000}
a:active {text-decoration;background-color: #CCCCCC; color: #006666}
-->
</style></head>

<body bgcolor="#CCCC99">
<div align="center"> 
<%if Request.ServerVariables("REQUEST_METHOD")<>"POST" then%>
  <form name="form1" action="recallpass.asp" method="POST">
    <p>��һ��</p>
    <p>�û����� 
      <input type="text" name="name">
      <input type="hidden" name="nameconfirm" value="ȷ��">
    </p>
    <p> 
      <input type="submit" name="confirm" value="ȷ��">
      <input type="reset" name="Submit2" value="��д">
    </p>
  </form>
<%end if%>  
<%if request("nameconfirm")="ȷ��" then
     set conn=server.CreateObject("ADODB.Connection")
     on error resume next
     conn.Open OpenString
     sql="select question,anwser from detail where username='"&request("name")&"'"
     set rs=conn.Execute(sql)
     for i=0 to conn.errors.count-1
        Response.Write conn.errors(i).description
        Response.Write conn.Errors(i).Number&"<br>"
     next
    ' Response.End
     if rs.eof then
       Response.Write "<div align=center>�Բ����û�����д������<a href='javascript:history.go(-1)'>����</a>��</div>"
       Response.End
     end if  
     session("user")=trim(request("name"))
     question=trim(rs("question"))
     session("anwser")=rs("anwser")
  %>
<form name="form2" action="recallpass.asp" method="POST">
    <p>�ڶ�������ش��������⣺<br><%=question%>?</p>
    <p>��ش�</p>
    <p>
      <input type="text" name="anwser">
      <input type="hidden" name="anwserconfirm" value="ȷ��">
    </p>
    <p> 
      <input type="submit" name="confirm" value="ȷ��">
      <input type="reset" name="Submit4" value="��д">
    </p>
  </form>
 <%end if
'Response.Write request("anwserconfirm")
 if request("anwserconfirm")="ȷ��" then
     'Response.Write "ok"
     if trim(request("anwser"))=trim(session("anwser")) then
        set conn=server.CreateObject("ADODB.Connection")
        conn.Open OpenString
        sql="select passwd from users where username='"&session("user")&"'"
 '       Response.Write sql
        set rs=conn.execute(sql)
        if rs.eof then
           Response.Write  "<div align=center>�Բ����û�����д������<a href='javascript:history.go(-2)'>����</a>��</div>"
           Response.End
        end if
        pass=trim(rs(0))
        session("user")=""%> 
  <p>�ش���ȷ�����������ǣ�<font color="red"><%=pass%></font></p>
   <%else%> 
  <p>�Բ������Ļش���ȷ������ȡ�����롣</p>
 <%end if%>
<p>   <a href="index.htm">����е�½</a></p>
<%end if%>
</div>
</body>
</html>
