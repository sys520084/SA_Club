<html>
<head>
<title>�û���½</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
body {
	background-color: #FFFFFF;
}
-->
</style>
</head>
<%session("user")=""%>
<body>
<div align="center">
  <form name="form1" action="jump.asp" method="post" >
    <p>&nbsp;</p>
    <table width="500" height="400" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="262"> 
          <p align="center">&nbsp;</p>
          <p align="center"><br>
            <br>
            <br>
            <br>
          </p>
          <p align="center">&nbsp; </p>
          <p align="center">�û����� 
            <input type="text" name="name" maxlength="20" size="15">
            ��� 
            <input type="password" name="pass" maxlength="20" size="15">
          </p>
          <p align="center"><a href="recallpass.asp">���������ˣ�</a> ��<a href="forum.htm">�οͲι�</a>����<a href="membership/memeber.asp?type=new">���û�</a></p>
          <p align="center"> 
            <input type="submit" name="Submit" value="ȷ��">
            <input type="reset" name="Submit2" value="��д">
            <br>
            <br>
            <br>
            </p>
          <p align="center">&nbsp; </p>        </td>
      </tr>
    </table>
    <p>&nbsp;</p>
  </form>
</div>
</body>
</html>