<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><!-- InstanceBegin template="/Templates/bs.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" --> 
<title>发送消息</title>
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
function move(loc)
{
	var desdir=window.showModalDialog(loc);
	//alert(desdir);
	if(desdir!= null){
	  return desdir;
	}
	return "";
}
-->
</script>
<style type="text/css">
<!--
.unnamed3 {
	font-family: "宋体";
	font-size: 9pt;
	text-decoration: underline blink;
	color: #000000;
}
-->
</style>
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
  <%
  if session("user")="" then
     Response.Write "<div align='center'>对不起，您需要先通过身份验证。</div>" 
     Response.End
  end if
  set conn=server.CreateObject("ADODB.Connection")
  Conn.open OpenString 
if not request("sendnow")="发送" then
  name1=session("user")
  name2=request("name1")
  id=request("id")
  retb=false
  tb=false
  if isnull(id) or id="" or isnull(title) then
     'Response.Write id&","&title
     tb=true
  end if   
  if not tb then
     sql="select title,sender,content,senddate,sendtime from message where id="&id
     set rs=conn.Execute(sql)
     if not rs.eof then
       title="回复："&trim(rs("title"))
       content1=">"&trim(rs("sender"))&"在"&trim(rs("senddate"))&" "&trim(rs("sendtime"))&"时说："
       contemp=">"&left(trim(rs("content")),100)&"..."
       if len(contemp)<>0 then 
       for i=1 to len(contemp)-1
         if asc(mid(contemp,i,1))=13 then
            contemp=left(contemp,i)&">"&right(contemp,len(contemp)-i-1)
         end if   
       next   
       end if
       content2=contemp
       retb=true
     end if  
  end if%>
  <p>请在这里编辑你想要发送的消息，然后点击发送按钮</p>
  <form method="post" action="sendmsg.asp">
    <table width="100%" height="234" border="1" cellpadding="1" cellspacing="1" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#6699FF">
      <tr> 
        <td width=143 bgcolor="#000066"><font color="#CCCCFF">发送者：</font></td>
        <td width="223"><font color="#333333"><%=name1%></font></td>
        <input type="hidden" name="sender" value="<%=name1%>">
        <td width=104 bgcolor="#000066"><font color="#CCCCFF">接收者：</font></td>
        <td width="276"><input name="reciever" value=<%=name2%>>
          <a href="#" class="unnamed3" onclick="javascript:reciever.value=move('userlist.asp');">选择接收者</a></td>
      </tr>
      <tr> 
        <td width=143 bgcolor="#000066"><font color="#CCCCFF">题目：</font></td>
        <td colspan="3"> <input type="text" name="title" size="60" <%if retb then%> value="<%=title%>"<%end if%>> 
        </td>
      </tr>
      <tr> 
        <td bgcolor="#000066" width="143"><font color="#CCCCFF">内容：</font></td>
        <td colspan="3" bgcolor="#6699FF"> </td>
      </tr>
      <tr bgcolor="#6699FF"> 
        <td height="188" colspan="4"> <div align="center"> 
            <textarea name="content" cols="80" rows="10"><%if retb then%><%=content1&chr(13)&content2%><%end if%></textarea>
            <br>
          </div></td>
      </tr>
    </table>
    <p> 
      <input type="submit" name="sendnow" value="发送">
      <input type="reset" name="Submit2" value="重添">
    </p>
  </form>
  <p>&nbsp;</p>
  <%end if
if request("sendnow")="发送" then
   title=replace(request("title"),"'","''")
   name1=request("sender")
   name2=request("reciever")
   content=replace(request("content"),"'","''")
   if title="" then
      Response.Write "<div align='center'>填写有错误：没写标题，<a href='javascript:history.go(-1)'>请重新填写</a></div>"
      Response.End
   end if
   if trim(name1)="" or trim(name2)="" then
      Response.Write "<div align='center'>错误发生：发送者或者接收者为空。<a href='javascript:history.go(-1)'>请重新填写</a></div>"
      Response.End   
   end if
   if trim(name1)=trim(name2) then
	  Response.Write "<div align='center'>错误发生：不能给自己发消息。<a href='javascript:history.go(-1)'>请重新填写</a></div>"
      Response.End   
   end if
   sql="insert into message (sender,reciever,title,senddate,sendtime,content,status,type) values ('"&_
       name1&"','"&name2&"','"&title&"','"&date&"','"&time&"','"&content&"',0,1)"
   conn.Execute(sql)
   %>
  <p>消息发送成功，请你<a href="msgman.asp">返回</a></p>
  <%end if%>
</div>
<!-- InstanceEndEditable -->
<p align="center"><font color="#CCFF66">集智俱乐部·版权所有 <font face="Arial, Helvetica, sans-serif"><br>
  Copyright 2002-2003 Clustering Intelligence Club</font></font> </p>
</body>
<!-- InstanceEnd --></html>
