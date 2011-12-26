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
<script language="javascript">
function delmsg(htmlurl){
  t=confirm("真的要删除这条消息吗？");
  if(t){
    window.location.replace(htmlurl);
  }
  return false;
}  
</script>
<title>消息查看</title></HEAD>
<BODY bgcolor="#FFFFFF">
<%if session("user")="" then
       Response.Write "<div align='center'>对不起，您需要先通过身份验证。</div>" 
       Response.End
  end if
  id=request("id")
  if trim(request("id"))="" then
       Response.Write "<div align='center'>对不起，您没有选中要察看的消息。</div>"  
       response.end
  end if
     set Conn=server.CreateObject("ADODB.Connection")
     Conn.open OpenString
     sql="select * from message where id="&id
     set rs=conn.Execute(sql)
     if rs.eof then
       Response.Write "<div align='center'>对不起，这条消息已经被删除了。</div>"
       Response.End 
     end if  
     if rs("status")=0 then
        sql="update message set status=1 where id="&id
        on error resume next
        conn.Errors.Clear 
        conn.Execute(sql)
        if conn.Errors.Count>0 then
           Response.Write("<div  align='center'>更新纪录时出错。</div>")
           conn.Errors.Clear 
        end if 
     end if%> 
<table width="100%" border="0" bgcolor="#CCCCCC">
  <tr> 
    <td width=60 bgcolor="#000066"><font color="#CCCCFF">发送者：</font></td>
    <td width="30%"><a href="../membership/display.asp?name=<%=trim(rs("sender"))%>"><%=trim(rs("sender"))%></a></td>
    <td width=60 bgcolor="#000066"><font color="#CCCCFF">接收者：</font></td>
    <td width="30%"><%=trim(rs("reciever"))%></td>
  </tr>
  <tr> 
    <td width=60 bgcolor="#000066"><font color="#CCCCFF">题目：</font></td>
    <td colspan="3"><%=trim(rs("title"))%></td>
  </tr>
  <tr> 
    <td width=60 bgcolor="#000066"><font color="#CCCCFF">发送时间：</font></td>
    <td colspan="3"><%=trim(rs("senddate"))&"&nbsp;"&trim(rs("sendtime"))%></td>
  </tr>
  <tr> 
    <td bgcolor="#000066"><font color="#CCCCFF">内容：</font></td>
    <td>
    </td>
    <td colspan="2"></td>
  </tr>
  <tr> 
    <td colspan="4" height="17" bgcolor="#CCCCFF"><%story=trim(rs("content"))
      if len(story)<>0 then 
       for i=1 to len(story)
         if asc(mid(story,i,1))=13 then
            story=left(story,i-1)&"<br>"&right(story,len(story)-i)
         end if   
       next   
       end if
       Response.Write story
       %> </td>
  </tr>
</table>
<p>对这条消息的操作：</p>
<P> 
  <input type="button" name="Button" value="回复" onclick="javascript:window.opener.location='sendmsg.asp?name1=<%=trim(rs("sender"))%>&id=<%=rs("id")%>';window.close();">
  <input type="button" name="Submit2" value="删除" onclick="javascript:delmsg('delmsg.asp?id=<%=rs("id")%>');">

</P>

</BODY>
</HTML>
