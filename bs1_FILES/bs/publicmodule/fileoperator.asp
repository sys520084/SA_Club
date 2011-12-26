<%
function UploadFile(MyPath,UserName,File_Max_Size,Max_FullSpace,fileTypes)
	dim Upload,file

	dim FilePath
	dim Count
	Set Upload = Server.CreateObject("Persits.Upload") 
	Upload.OverwriteFiles = True   '不能复盖
	Upload.IgnoreNoPost = True
	maxSz=FILE_MAX_SIZE
	Upload.SetMaxSize maxSz, True	 '限制大小

	Count = Upload.Save

	If Err.Number = 8 Then 
	   select case upload.err
		case 1
		Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		case 2
		Response.Write "文件大小超过了限制，总空间是"&MAX_FULLSPACE&"，请　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		end select
		UploadFile=""
		exit function
	Else 

	   If Err <> 0 Then 
		  Response.Write "错误信息: " & Err.Description 
	   Else
			If Count < 1 Then 
				Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
				UploadFile=false
			End If   
			i=0

			For Each file in Upload.Files	'列出所有上传文件
				'response.write i
				fileExt=lcase(replace(File.ext,".",""))
				'判断文件类型
				if lcase(fileEXT)="asp" and lcase(fileEXT)="asa" and lcase(fileEXT)="aspx" then
					response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
					response.end
				end if
				
				if fileTypes<>"" then
					atemp=split(fileTypes,",")
					pass=false
					for i=0 to ubound(atemp)
						if lcase(fileEXT)=atemp(i) then
							pass=true
							exit for
						end if
					next
					if not pass then
						response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
						response.end
					end if
				end if

				'文件变量付值
				address=MyPath
				file_name=UserName&"."&fileExt
				filename=address&file_name
				Filesize=File.Size
				'response.write filename
				'response.end
				file.saveas filename	'上传保存文件
				%>
				<p align="center">文件<font color="red"><%=file_name%></font>上传成功！</p>
				<%
				i=i+1
				set file=nothing
			Next
	   End If 
	End If
	set Upload =nothing
	UploadFile=file_name
end function
	%>

	<%
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
'	set upload=nothing %>
<!--	    <p>上传成功！</p>-->
	      <%
'	if filenames(1)="" and filenames(2)="" then
'		Response.Write "<div align=center>没有文件上传成功，<a href='javascript:history.go(-1);'>请返回</a></div>"
'		Response.End
'	end if
%>