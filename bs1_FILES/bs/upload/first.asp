<!--#include file="..\..\var.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><!-- InstanceBegin template="/Templates/bs.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>上传文件</title>
<!-- InstanceEndEditable --> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

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
<p align="center"></p>
<!-- InstanceBeginEditable name="content" --><div align=center>
<%'Response.End 
  session("loadtype")=""
  if session("user")="" then
         Response.Write "<div align='center'>对不起，您不能使用上传服务，可能是下列原因：<br>1.您的经验值不够，上传需要至少<font color=red>"&ALLOW_UP_VALUE&"</font>的经验值，请赶快发贴子挣取足够的经验值。<br>2.您今天已经超过了最大上传文件数目，清明天再试！<br>3.您的身份需要重新验证，请<a href='../membership/relogin.asp' onclick='return js_callpage(this.href)'>重新登陆</a>。</div>"
         Response.End
  end if
  'Response.Write request("type")
 
if request("type")<>"articles" then
  set conn=server.CreateObject("ADODB.Connection")
  set rs=server.CreateObject("ADODB.Recordset")
  conn.Open OpenString
  sql="select * from files where title=''"
  rs.Open sql,conn,3
   
  do while not rs.eof
	id=rs("id")
	filename=rs("filename")
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	'Set objFSO = Server.CreateObject("Scripting.FMG1TYKZ")
	if filename<>"" then
	 if objFSO.fileExists(WebAddress()&UPFILEPATH&filename) then
	      objFSO.DeleteFile(WebAddress()&UPFILEPATH&filename)
	 end if
	end if
	sql="delete from files where id="&id
	conn.Execute sql
	rs.MoveNext
 loop

 rs.Close 
	
  user=session("user")
  sql="select count(*) as cn from files where username='"&user&"' and uploaddate=#"&date&"#"
  sql1="select sum(bytes) as sm from files where username='"&user&"'"
  'Response.Write sql
  'Response.End 
  
  rs.Open sql,conn,3
  count=0
  size=0
  if not rs.EOF then
    count=rs("cn")
    rs.Close 
    rs.Open sql1,conn,3
    if not rs.EOF then
		if not isnull(rs("sm")) then
			size=rs("sm")
		end if
	end if
  end if
	rs.Close 
	'Response.Write "ok"&size
  %>
  <p><strong>用户<%=user%>上传文件管理</strong></p>
  <p>（每日上传文件数不超过<%=DAY_FILES%>个，最大总容量<font color=red><%=MAX_FULLSPACE/1000%> K</font>，您现在还有<font color=red><%=int((MAX_FULLSPACE-size)*100/1024)/100%> K</font>空间）</p>
  <%
     pageNo=request("page")
     PgSz=20

    sql="select * from files where username='"&user&"' order by id desc"
     rs.Open sql,conn,3	 
     rs.PageSize=cint(PgSz)
     all_page=rs.PageCount
     total=rs.RecordCount 
     if total<>0 then
         if pageNo<>"" then
            rs.AbsolutePage =pageNo
            RcoNo=1
         else
            rs.AbsolutePage =1
            rs.movefirst
            pageNo=1 
         end if%>
<div align="center"> 
  <%if pageNo>1 then%>
  <a href="first.asp?page=<%=pageNo-1%>">上一页</a>&nbsp; 
  <%end if%>
  <%if pageNo-all_page<0 or pageNo="" then%>
  <a href="first.asp?page=<%=pageNo+1%>">下一页</a> 
  <%end if%>
</div>

  <table width="100%" border="1" cellpadding="1" cellspacing="1" bordercolor="#6666FF">
    <tr align="center" bgcolor="#000033"> 
      <td width="13%">文件名</td>
      <td width="25%">文件标题</td>
      <td width="10%">上传日期</td>
      <td width="8%">下载次数</td>
      <td width="5%">大小K</td>
      <td width="10%">操作</td>
    </tr>
    <%for iPage=1 to rs.Pagesize
	          RecNo=(pageNo-1)*PgSz+iPage%>
    <tr bgcolor="#6666FF"> 
      <td align="center"><a href="download.asp?id=<%=rs("id")%>"><%=rs("filename")%></a></td>
      <td><font color="#000000"><a href="detail.asp?id=<%=rs("id")%>&url=first.asp&thispage=<%=pageNo%>"><%=rs("title")%></a></font></td>
      <td align="center"><font color="#000000"><%=rs("uploaddate")%></font></td>
      <td align="center"><font color="#000000"><%=rs("times")%></font></td>
      <td align="center"><font color="#000000"><%=int(rs("bytes")*100/1024)/100%></font></td>
      <td align="center"><font color="#000000"><a href="savedoc.asp?id=<%=rs("id")%>" onclick="return js_callpages(this.href)">编辑</a>&nbsp;
      <a href="delete.asp?id=<%=rs("id")%>" onclick="if(confirm('确实删除这个文件么？'))return js_callpages(this.href);">删除</a> </font></td>
      <%        rs.MoveNext
               if RecNo=total or rs.eof then
                 exit for
               end if%>
    </tr>
    <%next%>
  </table>
  <%   
     if pageNo<>"" then
        response.write "&nbsp;&nbsp;<font align=center> 第"&pageNo&"页/ 共"&all_page&"页</font>"
     else
        response.write "&nbsp;&nbsp;<font align=center> 第1页 / 共"&all_page&"页</font>"   
     end if
     rs.close%>
  &nbsp;&nbsp;查看第 
  <%for action=1 to all_page%>
  <a href="first.asp?page=<%=action%>">〖 <%=action%> 〗</a> 
  <%next%>
 <%else%>
  <div align=center>没有找到任何消息</div>

  <%end if
else
	session("loadtype")="articles"
end if 'end if type<>"articles"
  if count>=DAY_FILES then
	Response.Write "<div align=center>您已经超过了每天上传的最大文件数，不能再上传了！</div>"
  else 
	if size>=MAX_FULLSPACE then
		Response.Write "<div align=center>您已经超过了您所拥有的最大容量，不能再上传了！</div>"
	else
  %>
  
  <p><b>第一步：上传文件</b></p>
  <form method="post" name="form1" action="savedoc.asp" enctype="multipart/form-data">
    <p>请选择要上传的文件： 
      <input type="file" name="doc">
      （文件名尽量不用中文！） 
      <input type="submit" name="confirm" value="确定">
    </p>
    </form>
  <%end if
  end if
%>
</div><!-- InstanceEndEditable -->
<p align="center"><font color="#CCFF66">集智俱乐部·版权所有 <font face="Arial, Helvetica, sans-serif"><br>
  Copyright 2002-2003 Clustering Intelligence Club</font></font> </p>
</body>
<!-- InstanceEnd --></html>
