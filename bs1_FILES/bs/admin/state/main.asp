<!--#include file="../../../VAR.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>�߼���̳����</title>
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
<%catagory=request("catagory")
  if (not session("pass"&catagory)) and not session("admin") then
   Response.Write "<div align=center>�Բ��������ܴ򿪱�ҳ��</div>"
   Response.End
  end if
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  set rs=server.CreateObject("ADODB.Recordset")
  sql="select * from forum_state where catagory='"&catagory&"'"
  rs.Open sql,conn,3
  if rs.EOF then
     Response.Write "<div align=center>�Բ���û�ҵ������̳��</div>"
     Response.End
  end if
  %>   
<div align="center">
  <p>�߼���̳����:<font color=red><%=catagory%></font></p>
  <form name="form1" method="POST" action="modifystate.asp">
    <p>���ƻ�ӭ���֣� </p>
    <p> 
      <textarea name="welcome" cols="50" rows="5"><%=trim(rs("welcome"))%></textarea>
      <br>
    </p>
    <p>���ƽ������֣�</p>
    <p> 
      <textarea name="intro" cols="50" rows="5"><%=trim(rs("intro"))%></textarea>
    </p>
    <p> 
      <input type=hidden name="id" value=<%=rs("id")%>>
      <input type=hidden name="catagory" value="<%=catagory%>">
      <input type="submit" name="confirm" value="ȷ��">
      <input type="reset" name="reset" value="��д">
    </p>
    </form>
</div>
</body>
</html>
