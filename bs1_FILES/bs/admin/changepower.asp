<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:宋体;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
-->
</style>
<script language="javascript">
<!--
function js_callpages(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=330");
  newwin.focus();
  return false;
}
--> 
</script></HEAD>
<BODY bgcolor="#FFFFFF">
<%catagory=request("catagory")
if not session("pass"&catagory) and not session("admin") then
   Response.Write "<div align=center>对不起，您不能打开本页！</div>"
   Response.End
end if
if trim(Request("username"))="" then
   tb=true
end if
set conn=server.CreateObject("ADODB.Connection")
conn.Open OpenString
if not tb then
   username=request("username")
   sql="select * from power where catagory='"&catagory&"' and username='"&username&"' and power=1"
   set rs=conn.Execute(sql)
   if not rs.eof then
      Response.Write"<div align=center>此人的权限已被封闭了。</div>"
      tb=true
      rs.close
   else   
     sql="insert into power (username,catagory,power,fromdate) values ('"&username&"','"&catagory&"',1,'"&date&"')"
     'Response.Write sql
     'Response.End
     rs.close
     conn.Execute sql
   end if
     
end if   
   sql="select * from power where catagory='"&catagory&"' and power=1"
   if catagory="" then
      sql="select * from power where power=1"
   end if   
   set rs=conn.Execute(sql)
if not tb then   %> 
<P align=center>成功的封闭了：<%=username%>的权限。</P>
<%end if%>
<center><p>以下的用户权限仍被封闭着，点击它可以解除封闭：</p>
<p>
<%do while not rs.eof %>
     <a href="recoverpower.asp?id=<%=rs("id")%>&catagory=<%=trim(rs("catagory"))%>&username=<%=trim(rs("username"))%>" onclick="return js_callpages(this.href);"><%=trim(rs("username"))%></a>&nbsp;&nbsp;在<font color=red><%=trim(rs("catagory"))%></font>中的权限于<%=trim(rs("fromdate"))%>时被封<br>
    <% rs.movenext
  loop%>
  </p>
  <div align=center><a href="javascript:history.go(-1)">返回</a></div>
</center>     
</BODY>
</HTML>
