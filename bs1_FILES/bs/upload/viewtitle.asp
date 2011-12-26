<!--#include file="../../var.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>察看下载专区的主题</title>
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:"宋体";color:#333333;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0066FF}
a:visited {text-decoration: none;; color: #0033CC}
a:hover {text-decoration: underline; color: #0033CC}
a:active {text-decoration;color: #0066FF}
.art {
	font-size: 10pt;
	color: #0000CC;
	font-style: normal;
}
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
</head>
<body bgcolor="#003366" text="#CCCCCC" link="#0000CC" vlink="#000000">
<%function MyhtmlEncode(Str)
    dim result
    dim l
    l=len(str)
    result=""
	dim i
	for i = 1 to l
	    select case mid(str,i,1)
               case "'"
	               result=result+"''"
	           case else
	                result=result+mid(str,i,1)
	     end select
       next
       MyhtmlEncode= result
  end function
  for each j in Request.QueryString
       tail=tail&j&"="&request(j)&"&"
  next
  if right(tail,1)="&" then
     tail=left(tail,len(tail)-1)
  end if
  url="viewtitle.asp?"&tail
  session("address")=url
  set conn=server.CreateObject("ADODB.Connection")
  Conn.open OpenString
  isadmin=false
  'Response.Write catagory
  if session("passall") then
     isadmin=true
  end if     
  where=""
  if request("finder")="确定" then
     if not request("findtitle")="" then
        where=" and title like '%"&myhtmlencode(request("findtitle"))&"%'"
     end if
     if not request("findauthor")="" then   
       where=" and (username like '%"&myhtmlencode(request("findauthor"))&"%'"
     end if
     if not (request("fromdate")="" or request("todate")="") then
       where=" and (uploaddate between '"&request("fromdate")&"' and '"&request("todate")&"')"
     end if
  end if 
 
	order=request("order")
	
	if order="" then
		if session("order")="lastdate desc" then
			session("order")="id desc"
		end if

		order=session("order")
		if order="" then
			order="id desc"
		end if
	end if
	'Response.Write order
  session("order")=order
  session("page")="" 
  if trim(request("flag"))="ok" then
     'Response.Write request("page")
     sql=session("sql")
     sql=replace(sql,"order by id desc","order by "&order)
     sql=replace(sql,"order by times desc","order by "&order)
     page=cint(request("page"))
  else
	if not request("findauthor")="" then
		sql="select * from files where ture"&where&" order by " & order
	else
     	sql="select * from files where true"&where&" order by " & order 
	end if
     session("sql")=sql
  end if   
  'response.Write sql
  'response.end
  'on error resume next
  
	
  set rs=server.CreateObject("ADODB.Recordset")
 
  '删除所有的空标题的文件
  sql1="select * from files where title=''"
  rs.Open sql1,conn,3
  do while not rs.eof
	id=rs("id")
	filename=rs("filename")
	'Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFSO = Server.CreateObject("Scripting.FMG1TYKZ")
	if filename<>"" then
	 if objFSO.fileExists(WebAddress()&UPFILEPATH&filename) then
	      objFSO.DeleteFile(WebAddress()&UPFILEPATH&filename)
	 end if
	end if
	sql1="delete from files where id="&id
	conn.Execute sql1
	rs.MoveNext
  loop
  rs.Close
  
  if sql="" then
     sql="select * from files order"
  end if     
  'Response.Write sql
  'Response.End 
  rs.Open sql,conn,3
  pgSz=20
  if page="" then
     page=1
  end if   
  rs.PageSize=pgSz
  'Response.Write request("flag")
  session("page")=page
  if not rs.EOF then
      rs.AbsolutePage=page
  end if
  pgCount=rs.PageCount
  rsCount=rs.RecordCount
  if isnull(rsCount) then
     rsCount=0
  end if   
  if isnull(pgCount) then
     pgCount=0
  end if  
 
  %>
<form method="get" action="viewtitle.asp" name="frm1" onsubmit="if(frm1.page.value!='no'){frm1.submit();}">
  <table width="100%" border="0">
    <tr> 
      <td width="27%" valign="baseline"><font color="#00FFFF">排序方式</font><font color="#FFCC00">：
        <select name="order" id="order" onchange="frm1.submit();">
          <option value="id desc" <%if order="id desc" then%>selected<%end if%>>时间顺序</option>
          <option value="times desc" <%if order="times desc" then%>selected<%end if%>>下载次数</option>
        </select>
        </font></td>
       <td width="28%" height="19" valign="baseline"></td>
     <td width="45%" height="19" align="right">&nbsp;&nbsp; 
        <font color="#00FFFF">城市：下载专区</font> </td>
  </tr>
  </table>
  <table width="100%" border="1" bordercolor="#0099FF">
    <tr align="center"> 
      <td width="10%" height="24" bgcolor="#000066" class="art"><a href="../wolcome.asp?user=<%=session("user")%>" ><font color="#33FFFF">回到首页</font></a></td>
      <td width="10%" height="24" bgcolor="#7DD5FF"><a href="viewtitle.asp?catagory=<%=catagory%>&page=<%=page%>&flag=ok">刷新显示</a></td>
      <td width="11%" height="24" bgcolor="#7DD5FF"><a href="finder.htm" 
      >查找文件</a></td>
      <td width="12%" height="24" bgcolor="#7DD5FF">发表新文章</td>
      <td width="9%" height="24" bgcolor="#7DD5FF"> 
        <%if page>1 then%>
        <a href="viewtitle.asp?page=<%=page-1%>&flag=ok" > 
        <%end if%>
        上一页</a></td>
      <td width="7%" height="24" bgcolor="#7DD5FF"> 
        <%if page<pgCount then%>
        <a href="viewtitle.asp?page=<%=page+1%>&flag=ok"> 
        <%end if%>
        下一页</a></td>
      <td width="25%" height="24" bgcolor="#7DD5FF">第<%=page%>页/共<%=pgCount%>页,主题数<font color=red><%=rsCount%></font></td>
      <td width="16%" height="24" bgcolor="#7DD5FF">第 
        <input type="hidden" name="flag" value="ok">
        <select name="page" onchange="frm1.submit();">
          <%for i=1 to pgCount%> 
          <option value=<%=i%> <%if page=i then%> selected <%end if%>><%=i%></option>
          <%next%> 
        </select>
        页</td>
    </tr>
  </table>
 </form>
  <p></p>
  <% 
	if rs.EOF then
         Response.Write "<div align='center'><font color=white>对不起，没有找到任何可以下载的文件。</font></div>"
         Response.End
    end if%>   
<%if isadmin then%>
<form name=frm action="deal.asp" method="POST">
<%end if%>
  <table width="100%" border="0" cellpadding="0" cellspacing="1" bordercolorlight="#FFFFFF" bordercolordark="#000099" bgcolor="#FFFFFF"
  style="TABLE-LAYOUT: fixed; WORD-BREAK: break-all" >
    <tr align="center"> 
	  <%if isadmin then%><td width="5%" align="center" bgcolor="#99CCFF">&nbsp;</td><%end if%>
      <td width="15%" align="center" bgcolor="#99CCFF">上传者</td>
      <!--<td width="6%" bgcolor="#99CCFF">&nbsp;</td>-->
      <td width="25%" bgcolor="#99CCFF">题目</td>
      <td width="25%" bgcolor="#99CCFF">文件名</td>
      <td width="8%" bgcolor="#99CCFF">大小</td>
      <td width="6%" bgcolor="#99CCFF">点击率</td>
      <td width="12%" bgcolor="#99CCFF">上载时间</td>
    </tr>
    <% do while not rs.EOF and total<pgSz
         total=total+1
         id=rs("id")
         author=rs("username")%> 
   <tr align="center" <%if (total mod 2)=1 then%>bgcolor="#E1EAEA"<% else%>bgcolor="#ffffff"<%end if%>> 
	  <%if isadmin then%><td width="5%" align="center" bgcolor="#99CCFF"><input type="checkbox" name="deal" value=<%=id%>></td><%end if%>
      <td width="15%" align="left"  nowrap><a href="../membership/display.asp?name=<%=author%>" onClick="return js_callpages(this.href)"><%=author%></a></td>
      <!--<td width="6%"  align="left"><img src="../img/<%'=trim(rs("face"))&".gif"%>" align="absmiddle"></td>-->
      <td width="25%"  align="left"><a href="detail.asp?id=<%=id%>&thispage=<%=page%>&url=viewtitle.asp" ><%=trim(rs("title"))%></a> </td>
	  <td width="25%"  align="left"><a href="download.asp?id=<%=id%>"><%=rs("filename")%></a></td>
	  <td width="8%" ><%=int(rs("bytes")*100/1024)/100%>K</td>
      <td width="6%" ><%=rs("times")%></td>
      <td width="12%" ><%=trim(rs("uploaddate"))%></td>
    </tr>
    <%	rs.MoveNext
      loop%> 
  </table>
<%if isadmin then%>
<p></p>
<p></p>
<div align=center>
<input type="submit" value="确定删除">
<input type="reset" value="恢复">
</div>
</form>
<%end if%>

    <div align="center"> 
      <p><font color="#CCFF66">集智俱乐部·版权所有 <font face="Arial, Helvetica, sans-serif"><br>
        Copyright 2002-2003 Clustering Intelligence Club</font></font></p>
    </div>
</body>
</html>
