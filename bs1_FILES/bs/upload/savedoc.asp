<!--#include file="..\..\var.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><!-- InstanceBegin template="/Templates/bs.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>�ϴ��ļ�</title>
<!-- InstanceEndEditable --> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!-- InstanceBeginEditable name="head" -->
<script language="javascript">
<!--
function js_callpages(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=600,height=430");
  newwin.focus();
  return false;
}
-->
</script>
<body bgcolor="#043F80" link="#00FFFF" alink="#00FFFF" leftmargin="0" topmargin="0">
<style>

<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:white;}
a:link {text-decoration: none; color: yellow }
a:visited {
	text-decoration: none;
	color: #FFFF00;
}
a:hover {
	text-decoration: underline;
	color: #FF9900;
}
a:active {text-decoration;
	color: #FF9900;
}
.unnamed1 {  font-family: "����"; font-size: 9pt; font-style: normal; line-height: 100%; font-weight: normal; color: #99FFFF}
.unnamed2 {  font-family: "����"; font-size: 9pt; line-height: 100%; font-weight: normal; color: #294C39}
input {
	font-family: "����";
	font-size: 10pt;
	color: #FFFFFF;
	background-color: #000033;
}
area {
	font-family: "����";
	font-size: 9pt;
	color: #FFFFFF;
	background-color: #003366;
}
-->
</style>
<%'Response.End 
  if session("user")="" then
         Response.Write "<div align='center'>�Բ���������ʹ���ϴ�������<a href='../membership/relogin.asp' onclick='return js_callpage(this.href)'>���µ�½</a>��</div>"
         Response.End
  end if
if request("id")="" then%>

<!-- InstanceEndEditable --> 
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:white;}
a:link {text-decoration: none; color: yellow }
a:visited {
	text-decoration: none;
	color: #FFFF00;
}
a:hover {
	text-decoration: underline;
	color: #FF9900;
}
a:active {text-decoration;
	color: #FF9900;
}
.unnamed1 {  font-family: "����"; font-size: 9pt; font-style: normal; line-height: 100%; font-weight: normal; color: #99FFFF}
.unnamed2 {  font-family: "����"; font-size: 9pt; line-height: 100%; font-weight: normal; color: #294C39}
input {
	font-family: "����";
	font-size: 10pt;
	color: #FFFFFF;
	background-color: #000033;
}
textarea {
	font-family: "����";
	font-size: 9pt;
	color: #FFFFFF;
	background-color: #003366;
}
-->
</style>
</head>

<body bgcolor="#043F80" link="#00FFFF" alink="#00FFFF" leftmargin="0" topmargin="0">
<p align="center"><img src="../img/index.jpg" width="600" height="140"></p>
<!-- InstanceBeginEditable name="content" -->
<div align=center>
  <div align="center"> 
<%    

	dim Upload,file
	dim FilePath
	dim Count
	Set Upload = Server.CreateObject("Persits.Upload")
	Upload.OverwriteFiles = false   '���ܸ���
	Upload.IgnoreNoPost = True
	maxSz=FILE_MAX_SIZE
	Upload.SetMaxSize maxSz, True	 '���ƴ�С
	address=WebAddress() & UPFILEPATH	
	'response.write address
	'response.end 
	Count = Upload.Save
'response.end 
	If Err.Number = 8 Then 
	   select case upload.err
		case 1
		Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		case 2
		Response.Write "�ļ���С������"&Request.QueryString("lastbytes")/1024*1024&"M���ƣ��ܿռ���"&MAX_FULLSPACE&"���롡[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		end select
		Response.End 
	Else 

	   If Err <> 0 Then 
		  Response.Write "������Ϣ: " & Err.Description 
	   Else
			If Count < 1 Then 
				Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				Response.End
			End If   
			i=0
			'response.write uploadweb&"1ok<br>"
			'response.end

			For Each file in Upload.Files	'�г������ϴ��ļ�
				'response.write i
				fileExt=lcase(replace(File.ext,".",""))
				'�ж��ļ�����
				if lcase(fileEXT)="asp" and lcase(fileEXT)="asa" and lcase(fileEXT)="aspx" then
					response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
					response.end
				end if

				'�ļ�������ֵ
				address=WebAddress() & UPFILEPATH 
				'Response.write address
			 					
				'file_name=replace(replace(replace(now()," ",""),"-",""),":","") &File.ext
				uptime=replace(replace(replace(now()," ",""),"-",""),":","")
				'file_name=session("user")&uptime&"_"&lcase(file.FileName)
				file_name=session("user")&uptime&"."&lcase(fileExt)
				'response.write file_name
				'response.end
				rename=filename & "|"
				filename=address&file_name
				'response.write File.Size
				
				Filesize=File.Size
				
				'response.write filename
				'response.end
				file.saveas filename	'�ϴ������ļ�
				
				%>
				<p>�ϴ��ɹ���</p>
					<p>�ļ���<%=file_name%></p>
				<%
				i=i+1
				'filenames(i)=file_name
				set file=nothing
				
			Next

'response.end
	   End If 
	End If
	set Upload =nothing
'	if filenames(1)="" and filenames(2)="" then
'		Response.Write "<div align=center>û���ļ��ϴ��ɹ���<a href='javascript:history.go(-1);'>�뷵��</a></div>"
'		Response.End
'	end if
	
'Response.Write uploadweb&"ok"
'Response.End

'	Server.ScriptTimeOut=999999'Ҫ�������̳֧���ϴ����ļ��Ƚϴ󣬾ͱ������á�
	'�������
'	dim Forumupload,ranNum
'	dim formName,formPath,filename,file_name,fileExt,Filesize,F_Type
'	dim upNum,dateupnum
'	dim rename,DownloadID
'	dim filenames(2)
'	dim upload,file
'	set upload=new UpFile_Class ''�����ϴ�����
'	upload.GetDate FILE_MAX_SIZE   '��������-1��ʾ���޴�С
'
'	if upload.err > 0 then
'	    select case upload.err
'		case 1
'		Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
'		case 2
'		Response.Write "�ļ���С������"&Request.QueryString("lastbytes")/1024*1024&"M���ƣ��ܿռ���"&MAX_FULLSPACE&"���롡[ <a href=# 'onclick=history.go(-1)>�����ϴ�</a> ]"
'		end select
'		Response.End 
'	else
'		'formPath=upload.form("doc")
'		'��Ŀ¼���(/)
'		'if right(formPath,1)<>"/" then formPath=formPath&"/"
'		i=0
'		for each formName in upload.file ''�г������ϴ��˵��ļ�
'			set file=upload.file(formName)  ''����һ���ļ�����
'			fileExt=lcase(file.FileExt)
'			file_name=session("user")&"_"&lcase(file.FileName)
'			'�ж��ļ�����
'			if lcase(fileEXT)="asp" and lcase(fileEXT)="asa" and lcase(fileEXT)="aspx" then
'				response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
'				response.end
'			end if
'		
'			'��ֵ����
'			address=WebAddress() & UPFILEPATH 
'			FilePath= address & file_name
'			Filesize=file.FileSize
'			'if Filesize>5242880 then
'			'	Response.Write "�ļ���С���������� 5M��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
'			'	Response.End
'			'end if	
'			'��¼�ļ�
'			if Filesize>0 then         '��� FileSize > 0 ˵�����ļ�����
'				i=i+1
'				if file_name="" then
'					Response.Write "<div align>�Բ������ǲ�֧�������ļ����ϴ������ȸ����ļ�����<a href='javascript:history.go(-1)'>�뷵��</a>"
'					Response.End
'					'file_name="chinease1"&unk
'				end if
'
'				if i>1 then
'					Response.Write "<div algin=center>�ϴ��ļ�����Ŀ����1���ˣ���������<a href='javascript:history.go(-1)'>�뷵��</a></div>"
'					Response.End
'				end if
'				'Response.Write FilePath
'				'Set MagicFs = Server.CreateObject("Scripting.FileSystemObject")
'				Set MagicFs = Server.CreateObject("Scripting.FMG1TYKZ")
'				if MagicFs.FileExists(FilePath) then
'					Response.Write "<div align=center>�Ѿ������ļ���Ϊ��<font color=red>" & FilePath &"</font>���ļ���<a 'href='javascript:history.go(-1)'>���������������ϵ��ļ����ϴ���</a>"
'					file_name=""
'					Response.End
'				end if
'				file.SaveToFile FilePath   ''ִ���ϴ��ļ�
'				filenames(i)=file_name
'			end if
'			set file=nothing
'		next
'	end if
'	set upload=nothing

'	if filenames(1)="" and filenames(2)="" then
'		Response.Write "<div align=center>û���ļ��ϴ��ɹ���<a href='javascript:history.go(-1);'>�뷵��</a></div>"
'		Response.End
'	end if

	set conn=server.CreateObject("ADODB.Connection")
	conn.Open OpenString
	set rs=server.CreateObject("ADODB.Recordset")
	
	rs.Open "files",conn,1,3
	rs.AddNew 
	rs("username")=session("user")
	rs("filename")=file_name
	rs("extendname")=fileExt
	rs("uploaddate")=date
	rs("bytes")=clng(filesize)
	rs("title")=""
	rs.Update
	rs.Close
	user=session("user")

else
	sql="select * from files where id="&request("id")
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open OpenString
	set rs=server.CreateObject("ADODB.Recordset")
	rs.Open sql,conn,1,3
	if rs.eof then
		Response.Write "<div align=center>û���ҵ�Ҫ�޸ĵ��ļ�</div>"
		Response.End
	end if
	file_name=rs("filename")
	fileExt=rs("extendname")
	user=rs("username")
	filesize=rs("bytes")
	title=rs("title")
	des=rs("des")
	author=rs("author")
	journal=rs("journal")
	publishtime=rs("publishtime")
	keywords=rs("keywords")
	rs.close
end if

%>
    <p align=center>���ϴ����ļ����ӱ�����������������Ϊ�����ļ��ᱻ�Զ�ɾ��</p>
    <p></p>
    <form name="form1" method="post" action="savedes.asp" onSubmit="javascript:if(title.value==''){confirm('�ļ����ⲻ��Ϊ��');title.getfocus();}">
      <table width="100%" border="1" cellpadding="1" cellspacing="1" bordercolor="#000033" bgcolor="#6666FF">
        <tr> 
          <td bgcolor="#000033">�ļ�����</td>
          <td><%=file_name%> <input name="filename" type="hidden" value="<%=file_name%>" id="filename3"></td>
        </tr>
        <tr> 
          <td bgcolor="#000033">��չ����</td>
          <td><%=fileExt%> <input name="extname" type="hidden" value=<%=fileExt%> id="extname2"></td>
        </tr>
        <tr> 
          <td bgcolor="#000033">�ϴ��û���</td>
          <td><%=user%><input name="user" type="hidden" value=<%=user%>></td>
        </tr>
        <tr> 
          <td bgcolor="#000033">�ļ���С��</td>
          <td><%=filesize%> <input name="bytes" type="hidden" value=<%=filesize%> id="bytes"></td>
        </tr>
        <tr>
          <td bgcolor="#000033">���⣺</td>
          <td><input name="title" type="text" id="title" value="<%=title%>" size="50"></td>
        </tr>
        <tr> 
          <td bgcolor="#000033">�ļ�����/����ժҪ��</td>
          <td><textarea name="des" cols="60" rows="4" id="textarea"><%=des%></textarea></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">�����ƪ������ѧ�����ģ���ô���ṩ������Ϣ��</td>
          <td></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">�ؼ��֣�</td>
          <td><textarea name="keywords" cols="60" rows="4" id="textarea"><%=keywords%></textarea></td>
        </tr>
		<tr> 
		<tr> 
          <td bgcolor="#000033">�������ߣ�</td>
          <td><textarea name="author" cols="60" rows="3" id="author"><%=author%></textarea></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">��������/�ļ���</td>
          <td><input name="journal" type="text" id="journal"><%=journal%></textarea></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">����ʱ�䣺</td>
          <td><input name="publishtime" type="text" id="publishtime" value="<%=publishtime%>"></td>
        </tr>
      </table>
      <p> 
        <input type="submit" name="Submit" value="�����ļ����������ݿ�">
        <input name="re" type="reset" id="re" value="��д">
		<input type="hidden" name="uptype" value=<%response.write Request.QueryString("uptype")%>>
      </p>
      <p>&nbsp; </p>
    </form>
  </div>
</div>
<!-- InstanceEndEditable -->
<p align="center"><font color="#CCFF66">���Ǿ��ֲ�����Ȩ���� <font face="Arial, Helvetica, sans-serif"><br>
  Copyright 2002-2003 Clustering Intelligence Club</font></font> </p>
</body>
<!-- InstanceEnd --></html>