<!--#include file="..\..\var.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><!-- InstanceBegin template="/Templates/bs.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>上传文件</title>
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
.unnamed1 {  font-family: "宋体"; font-size: 9pt; font-style: normal; line-height: 100%; font-weight: normal; color: #99FFFF}
.unnamed2 {  font-family: "宋体"; font-size: 9pt; line-height: 100%; font-weight: normal; color: #294C39}
input {
	font-family: "宋体";
	font-size: 10pt;
	color: #FFFFFF;
	background-color: #000033;
}
area {
	font-family: "宋体";
	font-size: 9pt;
	color: #FFFFFF;
	background-color: #003366;
}
-->
</style>
<%'Response.End 
  if session("user")="" then
         Response.Write "<div align='center'>对不起，您不能使用上传服务，请<a href='../membership/relogin.asp' onclick='return js_callpage(this.href)'>重新登陆</a>。</div>"
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
.unnamed1 {  font-family: "宋体"; font-size: 9pt; font-style: normal; line-height: 100%; font-weight: normal; color: #99FFFF}
.unnamed2 {  font-family: "宋体"; font-size: 9pt; line-height: 100%; font-weight: normal; color: #294C39}
input {
	font-family: "宋体";
	font-size: 10pt;
	color: #FFFFFF;
	background-color: #000033;
}
textarea {
	font-family: "宋体";
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
	Upload.OverwriteFiles = false   '不能复盖
	Upload.IgnoreNoPost = True
	maxSz=FILE_MAX_SIZE
	Upload.SetMaxSize maxSz, True	 '限制大小
	address=WebAddress() & UPFILEPATH	
	'response.write address
	'response.end 
	Count = Upload.Save
'response.end 
	If Err.Number = 8 Then 
	   select case upload.err
		case 1
		Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		case 2
		Response.Write "文件大小超过了"&Request.QueryString("lastbytes")/1024*1024&"M限制，总空间是"&MAX_FULLSPACE&"，请　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		end select
		Response.End 
	Else 

	   If Err <> 0 Then 
		  Response.Write "错误信息: " & Err.Description 
	   Else
			If Count < 1 Then 
				Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
				Response.End
			End If   
			i=0
			'response.write uploadweb&"1ok<br>"
			'response.end

			For Each file in Upload.Files	'列出所有上传文件
				'response.write i
				fileExt=lcase(replace(File.ext,".",""))
				'判断文件类型
				if lcase(fileEXT)="asp" and lcase(fileEXT)="asa" and lcase(fileEXT)="aspx" then
					response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
					response.end
				end if

				'文件变量付值
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
				file.saveas filename	'上传保存文件
				
				%>
				<p>上传成功！</p>
					<p>文件名<%=file_name%></p>
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
'		Response.Write "<div align=center>没有文件上传成功，<a href='javascript:history.go(-1);'>请返回</a></div>"
'		Response.End
'	end if
	
'Response.Write uploadweb&"ok"
'Response.End

'	Server.ScriptTimeOut=999999'要是你的论坛支持上传的文件比较大，就必须设置。
	'定义变量
'	dim Forumupload,ranNum
'	dim formName,formPath,filename,file_name,fileExt,Filesize,F_Type
'	dim upNum,dateupnum
'	dim rename,DownloadID
'	dim filenames(2)
'	dim upload,file
'	set upload=new UpFile_Class ''建立上传对象
'	upload.GetDate FILE_MAX_SIZE   '这里输入-1表示不限大小
'
'	if upload.err > 0 then
'	    select case upload.err
'		case 1
'		Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
'		case 2
'		Response.Write "文件大小超过了"&Request.QueryString("lastbytes")/1024*1024&"M限制，总空间是"&MAX_FULLSPACE&"，请　[ <a href=# 'onclick=history.go(-1)>重新上传</a> ]"
'		end select
'		Response.End 
'	else
'		'formPath=upload.form("doc")
'		'在目录后加(/)
'		'if right(formPath,1)<>"/" then formPath=formPath&"/"
'		i=0
'		for each formName in upload.file ''列出所有上传了的文件
'			set file=upload.file(formName)  ''生成一个文件对象
'			fileExt=lcase(file.FileExt)
'			file_name=session("user")&"_"&lcase(file.FileName)
'			'判断文件类型
'			if lcase(fileEXT)="asp" and lcase(fileEXT)="asa" and lcase(fileEXT)="aspx" then
'				response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
'				response.end
'			end if
'		
'			'付值变量
'			address=WebAddress() & UPFILEPATH 
'			FilePath= address & file_name
'			Filesize=file.FileSize
'			'if Filesize>5242880 then
'			'	Response.Write "文件大小超过了限制 5M　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
'			'	Response.End
'			'end if	
'			'记录文件
'			if Filesize>0 then         '如果 FileSize > 0 说明有文件数据
'				i=i+1
'				if file_name="" then
'					Response.Write "<div align>对不起，我们不支持中文文件名上传，请先更改文件名！<a href='javascript:history.go(-1)'>请返回</a>"
'					Response.End
'					'file_name="chinease1"&unk
'				end if
'
'				if i>1 then
'					Response.Write "<div algin=center>上传文件的数目超过1个了，不合理！<a href='javascript:history.go(-1)'>请返回</a></div>"
'					Response.End
'				end if
'				'Response.Write FilePath
'				'Set MagicFs = Server.CreateObject("Scripting.FileSystemObject")
'				Set MagicFs = Server.CreateObject("Scripting.FMG1TYKZ")
'				if MagicFs.FileExists(FilePath) then
'					Response.Write "<div align=center>已经存在文件名为：<font color=red>" & FilePath &"</font>的文件，<a 'href='javascript:history.go(-1)'>请重新命名磁盘上的文件再上传！</a>"
'					file_name=""
'					Response.End
'				end if
'				file.SaveToFile FilePath   ''执行上传文件
'				filenames(i)=file_name
'			end if
'			set file=nothing
'		next
'	end if
'	set upload=nothing

'	if filenames(1)="" and filenames(2)="" then
'		Response.Write "<div align=center>没有文件上传成功，<a href='javascript:history.go(-1);'>请返回</a></div>"
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
		Response.Write "<div align=center>没有找到要修改的文件</div>"
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
    <p align=center>给上传的文件添加标题和描述，如果标题为空则文件会被自动删除</p>
    <p></p>
    <form name="form1" method="post" action="savedes.asp" onSubmit="javascript:if(title.value==''){confirm('文件标题不能为空');title.getfocus();}">
      <table width="100%" border="1" cellpadding="1" cellspacing="1" bordercolor="#000033" bgcolor="#6666FF">
        <tr> 
          <td bgcolor="#000033">文件名：</td>
          <td><%=file_name%> <input name="filename" type="hidden" value="<%=file_name%>" id="filename3"></td>
        </tr>
        <tr> 
          <td bgcolor="#000033">扩展名：</td>
          <td><%=fileExt%> <input name="extname" type="hidden" value=<%=fileExt%> id="extname2"></td>
        </tr>
        <tr> 
          <td bgcolor="#000033">上传用户：</td>
          <td><%=user%><input name="user" type="hidden" value=<%=user%>></td>
        </tr>
        <tr> 
          <td bgcolor="#000033">文件大小：</td>
          <td><%=filesize%> <input name="bytes" type="hidden" value=<%=filesize%> id="bytes"></td>
        </tr>
        <tr>
          <td bgcolor="#000033">标题：</td>
          <td><input name="title" type="text" id="title" value="<%=title%>" size="50"></td>
        </tr>
        <tr> 
          <td bgcolor="#000033">文件描述/文章摘要：</td>
          <td><textarea name="des" cols="60" rows="4" id="textarea"><%=des%></textarea></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">如果这篇文章是学术论文，那么请提供下列信息：</td>
          <td></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">关键字：</td>
          <td><textarea name="keywords" cols="60" rows="4" id="textarea"><%=keywords%></textarea></td>
        </tr>
		<tr> 
		<tr> 
          <td bgcolor="#000033">文章作者：</td>
          <td><textarea name="author" cols="60" rows="3" id="author"><%=author%></textarea></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">发表刊物/文集：</td>
          <td><input name="journal" type="text" id="journal"><%=journal%></textarea></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">发表时间：</td>
          <td><input name="publishtime" type="text" id="publishtime" value="<%=publishtime%>"></td>
        </tr>
      </table>
      <p> 
        <input type="submit" name="Submit" value="保存文件描述到数据库">
        <input name="re" type="reset" id="re" value="重写">
		<input type="hidden" name="uptype" value=<%response.write Request.QueryString("uptype")%>>
      </p>
      <p>&nbsp; </p>
    </form>
  </div>
</div>
<!-- InstanceEndEditable -->
<p align="center"><font color="#CCFF66">集智俱乐部·版权所有 <font face="Arial, Helvetica, sans-serif"><br>
  Copyright 2002-2003 Clustering Intelligence Club</font></font> </p>
</body>
<!-- InstanceEnd --></html>
