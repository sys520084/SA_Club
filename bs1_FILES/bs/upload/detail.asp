<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>������ϸ��Ϣ�鿴</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li,body,tr{font-size:9pt; line-height:14pt; font-family:??;color:#e7e7e7;}
a:link {
	text-decoration: none;
	color: #FFCC00;
}
a:visited {
	text-decoration: none;
	color: #FFCC00;
}
a:hover {text-decoration: underline; color: #0099FF}
a:active {text-decoration;
	color: #FFCC00;
}
.unnamed1 {  font-family: "����"; font-size: 9pt; font-weight: normal; color: #F7F7F7}
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
</script>
<style type="text/css">
<!--
.content {
	font-family: "����";
	font-size: 10pt;
	font-style: normal;
	line-height: 20px;
}
-->
</style>
</head>

<body bgcolor="#003366" text="#FFFFFF">
<%
id=0
  id=request("id")
  
     url=request("url")&"?page="&request("thispage")
     user=trim(session("user"))
     set conn=server.CreateObject("ADODB.Connection")
     conn.Open OpenString
     'response.write sql
     'response.end
     'Response.Write page
     set rs=server.CreateObject("ADODB.Recordset")
     'catagory=session("catagory")
     if session("passall") then
        'Response.Write "ok"
        isadmin=true
     end if
     sql="select * from files where id="&id
     session("address")="../upload/detail.asp?id="&id&"&thispage="&request("thispage")&"&url="&request("url")
     'Response.Write sql
     'Response.End 
     rs.Open sql,conn,3
     if rs.EOF then
        Response.Write "<div align='center'>�Բ�����Ҫ�鿴���ļ�û���ҵ������������Ѿ���ɾ���ˡ�</div>"
        Response.End
     end if
     'Response.Write page
     'on error resume next
%>
  <table width="100%" border="0">
    <tr> 
      <td width="72%">�ļ����⣺<%=trim(rs("title"))%></td>
      <td align="right" width="28%"><%if pic<>"" then%>
        <%end if%></td>
    </tr>
  </table>
  <table width="98%" border="0">
    <tr align="center" bgcolor="#003399"> 
      <td width="10%" height="22" bgcolor="#003399"><a href=<%=url%> class="unnamed1">����ҳ</a></td>
      <td width="12%" height="22"><a href="detail.asp?id=<%=id%>" class="unnamed1">ˢ����ʾ</a></td>
    </tr>
  </table>
<%j=0
  do while not rs.EOF%>
<table width="98%" border="0">
  <tr> <%
        title=server.URLEncode(left(trim(rs("title")),100))
       content=rs("des")
       username=trim(rs("username"))
       if len(content)>0 then
		  contemp=">"&trim(username)&":"&left(trim(content),100)&"..."
		  if len(contemp)<>0 then 
		  for i=1 to len(contemp)-1
		    if asc(mid(contemp,i,1))=13 then
		       contemp=left(contemp,i)&">"&trim(username)&":"&right(contemp,len(contemp)-i-1)
		    end if   
		  next   
		  end if
		  lastcontent=contemp
		end if  
       %> 
    <td colspan="2"> <a href="download.asp?id=<%=rs("id")%>" target="_blank">[�����ļ�]</a>&nbsp;&nbsp; 
      <a href="../forum/new.asp?catagory=����ר��&title=<%=title%>&flag=ok&id=<%=id%>&file=1&lastcontent=<%=lastcontent%>">[�ظ�����]</a> 
      &nbsp; 
      <%session("catagory")="����ר��"
		catagory=session("catagory")
    if session("user")=username or isadmin then%>
	  <a href="savedoc.asp?id=<%=rs("id")%>" onclick="return js_callpages(this.href)">[�༭�ļ����������]</a>&nbsp;&nbsp; 
      <a href="delete.asp?id=<%=rs("id")%>" onclick="if(confirm('ȷʵɾ������ļ�ô��'))return js_callpages(this.href);">[ɾ�����ļ�]</a>&nbsp;&nbsp; 
      <%end if%>
    </td>
  </tr>
  <tr> 
    <td colspan="2" height="18">
   <b><a href="../membership/display.asp?name=<%=username%>" onclick="return js_callpages(this.href)"><%=username%></a></b>��<%=trim(rs("uploaddate"))%>ʱ�ϴ����ļ���<font color="#CCFF00"><b><%=trim(rs("title"))%></b></font>��<%=int(rs("bytes")*100/1024)/100%>K��������</td>
  </tr>
  <tr align="center" valign="top"> 
    <td colspan="2" bgcolor="#006699" height="23"> 
      <table width="98%" border="0">
        <tr>
          <td class="content"><%
       content=rs("des")
       content=ViewImg(content)
       if content="" then
          content="&nbsp;"
       end if   
       Response.Write content
       %></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td height=2 bgcolor="#CCCCCC"></td>
  </tr>
</table>
<%
fatherid=rs("id")
rs.movenext
loop
'response.end
rs.close
	sql="select * from articles where fatherid=-"&fatherid
	rs.open sql,conn,3
	rsCount=rs.recordcount
	do while not rs.eof
	j=j+1%>
<table width="98%" border="0">
  <tr> 
    <%
     if left(trim(rs("title")),3)<>"�ظ���" then
        title=server.URLEncode("�ظ���"&left(trim(rs("title")),97))
     else
        title=server.URLEncode(left(trim(rs("title")),100))
     end if   
     if len(rs("content"))>0 then
       contemp=">"&trim(rs("author"))&":"&left(trim(rs("content")),100)&"..."
       if len(contemp)<>0 then 
       for i=1 to len(contemp)-1
         if asc(mid(contemp,i,1))=13 then
            contemp=left(contemp,i)&">"&trim(rs("author"))&":"&right(contemp,len(contemp)-i-1)
         end if   
       next   
       end if
       lastcontent=contemp
     end if  
       lastcontent=server.URLEncode(lastcontent)
       author=trim(rs("author"))
       if not author="" then
                 temp=split(author,"(")
                 username=temp(0)
       end if  %>
    <td colspan="2"> <a href="../forum/new.asp?catagory=<%=catagory%>&title=<%=title%>&flag=ok&id=<%=id%>&file=1&lastcontent=<%=lastcontent%>" >[�ظ�����]</a>&nbsp;&nbsp; 
      <%s=rsCount
    t=j
    'Response.Write s&","&t
    if t=s and page=pageCount and session("user")=username or isadmin then%>
      <a href="../forum/edit.asp?id=<%=rs("id")%><%if isadmin then%>&admin=ok<%end if%>&catagory=<%=catagory%>">[�༭����]</a>&nbsp;&nbsp; 
      <a href="../forum/deal.asp?id=<%=rs("id")%>&action=del&catagory=<%=catagory%>" onclick="return confirm('ȷʵɾ����ƪ����ô��')">[ɾ������]</a>&nbsp;&nbsp; 
      <%end if
     'end if
    'if isadmin and (rs("score")=0 or isnull(rs("score"))) then%>
    </td>
  </tr>
  <tr> 
    <td colspan="2" height="18"> 
      <img src="../img/<%=trim(rs("face"))%>.gif"> <b><a href="../membership/display.asp?name=<%=username%>" onclick="return js_callpages(this.href)"><%=author%></a></b>��<%=trim(rs("ondate"))&"&nbsp;"&trim(rs("ontime"))%>ʱ�ڴ�����<font color="#CCFF00"><b><%=trim(rs("title"))%></b></font>���ᵽ��</td>
  </tr>
  <tr align="center" valign="top"> 
    <td colspan="2" bgcolor="#006699" height="23"> <table width="98%" border="0">
        <tr> 
          <td class="content">
            <%
       content=rs("content")
       content=ViewImg(content)
       if content="" then
          content="&nbsp;"
       end if   
       Response.Write content
       %>
          </td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height=2 bgcolor="#CCCCCC"></td>
  </tr>
</table>
<p> 
  <%rs.MoveNext
  'Response.Write j
  loop
  rs.Close
  set rs=nothing%>
</p>
</body>
</html>
