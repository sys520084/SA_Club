<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
</head>

<body>
<form method="post" name="form1" action="test.asp" enctype="multipart/form-data" target="_blank">
    ��Ӹ����� 
      <input type="file" name="doc">
      <input type="submit" name="confirm" value="ȷ��">
    </form>
<%

if request.ServerVariables("REQUEST_METHOD")="POST" then
response.write "OOO"
Set Upload = Server.CreateObject("Persits.Upload")
Upload.OverwriteFiles = false   '���ܸ���
Upload.IgnoreNoPost = True
maxSz=10485760
Upload.SetMaxSize maxSz, True	 '���ƴ�С
'address=WebAddress() & UPFILEPATH	
	'response.write address
	'response.end 
	Count = Upload.Save
end if
%>
OOKKK
</body>
</html>
