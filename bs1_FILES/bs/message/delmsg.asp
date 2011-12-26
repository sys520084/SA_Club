<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<style>
<!--
div,td,p{font-size:9pt; line-height:14pt; font-family:宋体;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
-->
</style>
</HEAD>
<title>删除消息</title>
<BODY bgColor=white>


<P>&nbsp;</P>
<%if session("user")="" then
     Response.Write "<div align='center'>对不起，您需要先通过身份验证。</div>" 
     Response.End
  end if
  if request("id")="" then
     Response.Write "<div  align='center'>对不起，您没有选中要删除的记录。</div>"
     Response.End
  end if
  id=request("id")
  set conn=server.CreateObject("ADODB.Connection")
  Conn.open OpenString
  sql="delete from message where id="&id
  conn.Errors.Clear
  conn.Execute(sql)
  if conn.Errors.Count>0 then
     Response.Write "对不起，删除消息时出错，可能的原因是：这条消息已经不存在了。"
     conn.Errors.Clear 
     Response.End
  end if
%>
<p align="center">删除消息成功，如果您仍能看到这条消息，请按浏览器上的刷新按钮。</p>
<p align="center"><a href="javascript:window.close();">关闭此窗口</a></p>     
    
</BODY>
</HTML>
