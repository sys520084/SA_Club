<!--#include file="../publicmodule/function.asp"-->
<!--#include file="..\..\var.asp"-->
<html>
<head>
<title>�ϴ���Ƭ</title>
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
</STYLE></head>

<body bgcolor="#E2ECF5"><div align="left">
<div align="center"><strong><font color="#000099">�ϴ�ͷ��ͼƬ</font></strong></div>
<p><%if session("user")="" then%>
�Բ����û���Ϊ�գ������ϴ���Ƭ���������µ�½��
 <%Response.End
  end if
  portrait=request("portrait")%>
 </p><form name="imgUploader" id="imgUploader" action="saveimg.asp" method="post" enctype="multipart/form-data">
  <div align="center">
    <p><strong>���ڵ�ͷ��</strong><img src="<%=ShowPortrait(portrait)%>" align="middle"> 
    </p>
    <p> �ϴ��µ�ͼ�� 
      <input type="file" name="file">
    </p>
    <p>
      <input type="submit" name="Submit" value="Submit">
    </p>
  </div>
</form>
</body>
</html>
