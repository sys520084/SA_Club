<!--#include file="../VAR.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>�г�����˵��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:����;color:black;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0000CC}
a:visited {text-decoration: none;; color: #660066}
a:hover {text-decoration: underline; color: #FF0000}
a:active {text-decoration;background-color: #CCCCCC; color: #006666}
-->
</style>
</head>

<body bgcolor="#FFFFFF">
<%if Request.ServerVariables("REQUEST_METHOD")<>"POST" then%>
<div align="center"> 
  <p>�г���������</p>
</div>
<ul>
  <li>�������������������д����<a href="membership/memeber.asp?type=%22change%22" onclick="window.opener.location=this.href;window.close();">�Ǽ�����</a>��</li>
  <li>����д����������г������ɣ�</li>
</ul>
<p align="center"><%=session("user")%>Ҫ����<%=session("catagory")%>�е��г��������ǣ�</p>
<form name="apply" action="apply.asp" method="POST">
  <div align="center"> 
    <input type="hidden" name="user" value="<%=session("user")%>">
    <input type="hidden" name="catagory" value="<%=session("catagory")%>">
    <textarea name="reason" cols="50" rows="8"></textarea>
  </div>
  <p align="center"> 
    <input type="submit" name="Submit" value="ȷ��">
    <input type="reset" name="Submit2" value="��д">
  </p>
</form>
<%else
    set conn=server.CreateObject("ADODB.Connection")
    conn.Open OpenString
    set rs=server.CreateObject("ADODB.Recordset")
    rs.Open "application",conn,1,3
    rs.addnew
    rs("name")=trim(request("user"))
    rs("catagory")=trim(request("catagory"))
    rs("reason")=trim(request("reason"))
    rs("signdate")=date
    rs.Update
    rs.Close
    
    %>
    <div align=center>��������ɹ�����Ⱥ���š�</div>
<%end if%>    
</body>
</html>
 
