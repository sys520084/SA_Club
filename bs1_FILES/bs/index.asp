<html>
<head>
<title>用户登陆</title>
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
          <p align="center">用户名： 
            <input type="text" name="name" maxlength="20" size="15">
            口令： 
            <input type="password" name="pass" maxlength="20" size="15">
          </p>
          <p align="center"><a href="recallpass.asp">忘记密码了？</a> 　<a href="forum.htm">游客参观</a>　　<a href="membership/memeber.asp?type=new">新用户</a></p>
          <p align="center"> 
            <input type="submit" name="Submit" value="确定">
            <input type="reset" name="Submit2" value="重写">
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
