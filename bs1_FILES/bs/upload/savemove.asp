<!--#include file="..\..\var.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<!--#include file="../forum/aspFunctions.asp"-->

<%Dbpp=DB()
open1=DBBS()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
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
-->
</style>
</head>

<body bgcolor="#043F80" link="#00FFFF" alink="#00FFFF" leftmargin="0" topmargin="0">
<div align=center>
<%

  set conn=server.CreateObject("ADODB.Connection")
  conn.Open open1  
  isadmin=IsSuperAdmin(Conn,Session("user"))
    'response.end
  if not isadmin then
         Response.Write "<div align='center'>对不起，您不能使用这页</div>"
         Response.End
  end if
  	id=request("id")
	
  	root=server.MapPath("/")
	path1=replace(root & "/bs/files/","/","\")
	
	if request("mytype")="article" then
		path21= "/thesis/doc/"
	else
		path21= "/thesis/program/" 
	end if
	path2=root & replace(path21,"/","\")
	filename=replace(request("filename")," ","_")
	path1=path1 & filename
	path2=path2 & filename
	Set Upload = Server.CreateObject("Persits.Upload")
	'response.write request("filename") & "<br>"
	'response.write path1 & "," & path2
	'
	'Upload.CopyFile "d:\wwwroot\swarmagents.com\wwwroot\thesis\doc\ecal.pdf", "d:\wwwroot\swarmagents.com\wwwroot\thesis\doc\mmm.pdf"
	
	Upload.MoveFile path1, path2
	'response.end 
	response.write "文件移动成功！"
	'response.end
	'response.end
	'修改数据库记录
	sql="update files set filename='" & "http://www.swarmagents.cn" & path21 & filename &"' where id=" & id
	conn.execute sql
	conn.close
	response.write "修改数据库成功！"
	
	'更新thesis数据库
	conn.open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ="&Dbpp
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open "select * from thesis",conn,1,3
	rs.addnew
	rs("chi_title")=request("title")
	response.write request("author")
	'response.end 
	if not request("author")="" then
	 authors=split(request("author"),",")
	 if ubound(authors)<=1 then
	 	rs("author1")=request("author")
	 else 
	 	if ubound(authors)=2 then
	 		rs("author1")=authors(0)
			rs("author2")=authors(1)
			rs("author_other")=authors(2)
	 	else 
			response.write "ok"
	 		if ubound(authors)>2 then
	 			rs("author1")=authors(0)
				rs("author2")=authors(1)
				for i=2 to ubound(authors)
					authorsstr=authorsstr & authors(i) & ","
				next
				rs("author_other")=authorsstr
	  		end if
		end if
	  end if
	end if
	if request("mytype")="program" then
		rs("type")=1
	end if
	rs("chi_abstract")=request("abstract")
	rs("intro")=request("description")
	rs("chi_keywords")=request("keywords")
	rs("address_info")=1
	rs("publish_time")=request("publishtime")
	rs("publisher")=request("journal")
	'response.write request("filename")
	'response.end 
	rs("loc_doc")=request("filename")
	rs("uploader")=request("user")
	rs.update
	rs.close
		'conn.Close 
'response.write request("Tcatalog")
'response.end 
set rs1=server.CreateObject("ADODB.Recordset")
'Response.Write dbpath
'Response.End 
'Conn.open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DbPath
sql="Select top 1 *  From thesis where chi_title='"&replace(request("title"),"'","''")&"' order by id desc"
rs1.open sql,conn,3,2
  if rs1.EOF then
	Response.Write "插入文章失败，请<a href='javascript:history.back(1)'>重试!</a>"
    Response.end
  end if  
		Tid=trim(rs1("id"))
rs1.Close 

set rs2=server.CreateObject("ADODB.Recordset")
'Response.Write dbpath
'Response.End 
'Conn.open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DbPath
sql="Select * From relation"
rs2.open sql,conn,3,2
	'for i=0 to n-1
	a=split(request("Tcatalog"),",")
	for i=0 to ubound(a)-1
		rs2.addnew
		rs2("thesis_id")=Tid
		'Response.Write Tid		
		rs2("catalog_id")=a(i)
		'Response.Write rs2("catalog_id")
		rs2.update
	next
	'Response.End
	rs2.close
%>
<div align=center>文件上传成功，该窗口即刻自动关闭</div>
</div>
</body>
</html>
