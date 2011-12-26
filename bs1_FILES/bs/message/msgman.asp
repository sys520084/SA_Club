<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><!-- InstanceBegin template="/Templates/bs.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Untitled Document</title>
<!-- InstanceEndEditable --> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!-- InstanceBeginEditable name="head" -->
<script language="javascript">
<!--
function js_callpage(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=330");
  newwin.focus();
  return false;
}
function delmsg(htmlurl){
  t=confirm("真的要删除这条消息吗？");
  if(t){
    js_callpage(htmlurl);
  }
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
<p align="center"><img src="../img/index.jpg" width="600" height="140"></p>
<!-- InstanceBeginEditable name="content" --> 
<%'Response.Write session("user")

  if session("user")="" then
         Response.Write "<div align='center'>对不起，您不能使用消息服务，请<a href='../membership/relogin.asp' onclick='return js_callpage(this.href)'>重新登陆</a>。</div>"
         Response.End
  end if%>
<p align="center"><b><font color="#CC6600">消息管理器</font></b> </p>
<p align="center"><a href="sendmsg.asp">发送消息</a>&nbsp;<a href="saveall.asp" target="_blank">保存所有消息</a></p>
<% 
     name1=session("user")
     pageNo=request("page")
     PgSz=20
     maxMsgCount=MAX_NUM_MSG
     set Conn=server.CreateObject("ADODB.Connection")
     set rs=server.CreateObject("ADODB.Recordset")
     Conn.open OpenString
     sql="select top "&maxMsgCount&" * from message where reciever='"&name1&"' or (type=0 and reciever='all') order by id"
'     Response.Write sql
'	 Response.End 
     rs.Open sql,conn,3
     minid=0
     if not rs.EOF then
		rs.MoveLast 
		minid=rs("id")
	 end if
	 rs.Close 
	 'Response.Write minid
     sql="select * from message where (reciever='"&name1&"' or (type=0 and reciever='all')) and id<="&minid&" order by id desc"
     'Response.Write sql
	 'Response.End 
     rs.Open sql,conn,3
     rs.PageSize=cint(PgSz)
     all_page=rs.PageCount
     total=rs.RecordCount 
     if total<>0 then
		 if total>=maxMsgCount then
			Response.Write "<div align='center'>您的短消息已经超过<font color='ff00f0'>"&maxMsgCount&"</font>条，需要删除以前的才能读取新消息</div>"
		 else
		    response.write "<div align='center'>共有 <font color='ff00f0'>"&total&"</font> 条信息。</div>"
		 end if
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
  <a href="msgman.asp?page=<%=pageNo-1%>&order=<%=order%>">上一页</a>&nbsp; 
  <%end if%>
  <%if pageNo-all_page<0 or pageNo="" then%>
  <a href="msgman.asp?page=<%=pageNo+1%>&order=<%=order%>">下一页</a> 
  <%end if%>
</div>
<div align="center">
  <table width="90%" border="1" cellpadding="1" cellspacing="1" bordercolor="#0066FF" bordercolorlight="#666666" bordercolordark="#CCCCCC">
    <tr bgcolor="#000066" align="center"> 
      <td width="35%"><font color="#CCCCFF">题目</font></td>
      <td width="15%"><font color="#CCCCFF">发送者</font></td>
      <td width="20%"><font color="#CCCCFF">发送时间</font></td>
      <td width="15%"><font color="#CCCCFF">操作</font></td>
    </tr>
    <%for iPage=1 to rs.Pagesize
	          RecNo=(pageNo-1)*PgSz+iPage%>
    <tr bgcolor="<%if rs("type")=0 then%>#6699FF<%else%>#6666FF<%end if%>" align="center"> 
      <td width="50%"><a href="viewmsg.asp?id=<%=rs("id")%>" onclick="return js_callpage(this.href)"><%=trim(rs("title"))%></a> 
        <%if trim(rs("content"))="" or isnull(rs("content")) then%>
        <font color="#000000">（内空） </font>
<%end if%>
        &nbsp; 
        <%if rs("status")=0 then%>
        <font color="red">新！</font> 
        <%end if%>
      </td>
      <td width="15%"><a href="../membership/display.asp?name=<%=trim(rs("sender"))%>"  onclick="return js_callpage(this.href)"><%=trim(rs("sender"))%></a></td>
      <td width="20%"><%=trim(rs("senddate"))&"&nbsp;"&trim(rs("sendtime"))%></td>
      <td width="5%"> 
        <a href="sendmsg.asp?name1=<%=trim(rs("sender"))%>&id=<%=rs("id")%>&user=<%=session("user")%>"><img src="../img/re.gif" alt="回复发送者" width="14" height="17" border="0"></a><%if rs("type")=1 then%>&nbsp;<a href="delmsg.asp?id=<%=rs("id")%>&user=<%=session("user")%>" onclick="return delmsg(this.href);"><img src="../img/del.gif" alt="删除" width="15" height="15" border="0"><%end if%></a>
        <%     rs.MoveNext
               if RecNo=total or rs.eof then
                 exit for
               end if%>
    </tr>
    <%         
          next%>
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
  <a href="msgman.asp?page=<%=action%>&order=<%=order%>">〖 <%=action%> 〗</a> 
  <%next%>
  &nbsp;页 
  <% else%>
  <div align=center>没有找到任何消息</div>
  <%end if%>
</div>
<!-- InstanceEndEditable -->
<p align="center"><font color="#CCFF66">集智俱乐部·版权所有 <font face="Arial, Helvetica, sans-serif"><br>
  Copyright 2002-2003 Clustering Intelligence Club</font></font> </p>
</body>
<!-- InstanceEnd --></html>
