<%
function UploadFile(MyPath,UserName,File_Max_Size,Max_FullSpace,fileTypes)
	dim Upload,file

	dim FilePath
	dim Count
	Set Upload = Server.CreateObject("Persits.Upload") 
	Upload.OverwriteFiles = True   '���ܸ���
	Upload.IgnoreNoPost = True
	maxSz=FILE_MAX_SIZE
	Upload.SetMaxSize maxSz, True	 '���ƴ�С

	Count = Upload.Save

	If Err.Number = 8 Then 
	   select case upload.err
		case 1
		Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		case 2
		Response.Write "�ļ���С���������ƣ��ܿռ���"&MAX_FULLSPACE&"���롡[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		end select
		UploadFile=""
		exit function
	Else 

	   If Err <> 0 Then 
		  Response.Write "������Ϣ: " & Err.Description 
	   Else
			If Count < 1 Then 
				Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				UploadFile=false
			End If   
			i=0

			For Each file in Upload.Files	'�г������ϴ��ļ�
				'response.write i
				fileExt=lcase(replace(File.ext,".",""))
				'�ж��ļ�����
				if lcase(fileEXT)="asp" and lcase(fileEXT)="asa" and lcase(fileEXT)="aspx" then
					response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
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
						response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
						response.end
					end if
				end if

				'�ļ�������ֵ
				address=MyPath
				file_name=UserName&"."&fileExt
				filename=address&file_name
				Filesize=File.Size
				'response.write filename
				'response.end
				file.saveas filename	'�ϴ������ļ�
				%>
				<p align="center">�ļ�<font color="red"><%=file_name%></font>�ϴ��ɹ���</p>
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
'					Response.Write "<div algin=center>�ϴ��ļ�����Ŀ����1���ˣ�������<a href='javascript:history.go(-1)'>�뷵��</a></div>"
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
'	set upload=nothing %>
<!--	    <p>�ϴ��ɹ���</p>-->
	      <%
'	if filenames(1)="" and filenames(2)="" then
'		Response.Write "<div align=center>û���ļ��ϴ��ɹ���<a href='javascript:history.go(-1);'>�뷵��</a></div>"
'		Response.End
'	end if
%>