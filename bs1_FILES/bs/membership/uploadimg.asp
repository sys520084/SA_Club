<!--#include file="../publicmodule/function.asp"-->
<!--#include file="..\..\var.asp"-->
<html>
<head>
<title>上传照片</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>DIV {
	FONT-FAMILY: "宋体"; FONT-SIZE: 9pt
}
TD {
	FONT-FAMILY: "宋体"; FONT-SIZE: 9pt
}
A {
	COLOR: #003366; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none
}
A:hover {
	COLOR: #ff0000; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: underline
}
</STYLE></head>

<body bgcolor="#E2ECF5"><div align="left">
<div align="center"><strong><font color="#000099">上传头像图片</font></strong></div>
<p><%if session("user")="" then%>
对不起，用户名为空，不能上传照片，请您重新登陆！
 <%Response.End
  end if
  portrait=request("portrait")%>
 </p><form name="imgUploader" id="imgUploader" action="saveimg.asp" method="post" enctype="multipart/form-data">
  <div align="center">
    <p><strong>现在的头像：</strong><img src="<%=ShowPortrait(portrait)%>" align="middle"> 
    </p>
    <p> 上传新的图像： 
      <input type="file" name="file">
    </p>
    <p>
      <input type="submit" name="Submit" value="Submit">
    </p>
  </div>
</form>
</body>
</html>
