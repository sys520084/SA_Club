<!--#include file="../../VAR.asp"-->
<!--#include file="../forum/aspFunctions.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>察看主题</title><style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:宋体;color:black;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0000CC}
a:visited {text-decoration: none;; color: #660066}
a:hover {text-decoration: underline; color: #FF0000}
a:active {text-decoration;background-color: #CCCCCC; color: #006666}
-->
</style>
<script language="javascript">
<!--
function js_callpages(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=300");
  newwin.focus();
  return false;
}
function changevalue(txt){
  frmpower.username.value=txt;
}  
-->
</script>
</head>
<body bgcolor="#FFFFFF" link="#000000" vlink="#660000" onload="deleted=false;deletedall=false;">
<%
set conn=server.CreateObject("ADODB.Connection")
     Conn.open OpenString
     count=0
 catagory=request("catagory")
 isadmin=IsSuperAdmin(Conn,Session("user"))
'Response.Write catagory
  'Response.Write "pass"&catagory
   ' Response.Write session("pass"&catagory)

 if not isadmin then
      Response.Write "<div align=center>对不起，您不是本版的管理员，请退出</div>"
      Response.End
 end if
  conn.BeginTrans 
 '先把所有的主体性文章的Fatherid遍为它的id
  sql="update articles set fatherid=id where fatherid=0"
  conn.Execute sql
  set rs=server.CreateObject("ADODB.Recordset")
  sql="select * from articles where catagory='"&catagory&"' order by fatherid desc,id" 
  rs.Open sql,conn,3
  pgSz=20
  page=1
  rs.PageSize=pgSz
  'Response.Write request("flag")
  if trim(request("flag"))="ok" then
     'Response.Write request("page")
     page=cint(request("page"))
  end if
  if not rs.EOF then
      rs.AbsolutePage=page
  end if
  pgCount=rs.PageCount
  if isnull(pgCount) then
     pgCount=0
  end if   
  %> </p>
<p align="center">论坛管理： <%=catagory%></p>
<p align="center"><a href="elite/main.asp?catagory=<%=catagory%>">精华区整理</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="state/main.asp?catagory=<%=catagory%>" target="_blank">高级管理</a></p>
<form method="get" action="admin.asp" name="frm1" onsubmit="if(frm1.page.value!='no'){frm1.submit();}"> 
  <table width="100%" border="1" bgcolor="#0099FF" bordercolorlight="#003366" bordercolordark="#CCFFFF" cellpadding="0" cellspacing="0">
    <tr align="center"> 
      <td width="17%" bgcolor="#000066"><a href="../forum/viewtitle.asp?catagory=<%=catagory%>" ><font color="#CCCCFF">显示文章主题</font></a></td>
      <td width="13%"><a href="admin.asp?catagory=<%=catagory%>&page=<%=page%>&flag=ok">刷新显示</a></td>
      <td width="10%"><%if page>1 then%><a href="admin.asp?page=<%=page-1%>&flag=ok&catagory=<%=catagory%>" ><%end if%>上一页</a></td>
      <td width="9%"><%if page<pgCount then%><a href="admin.asp?page=<%=page+1%>&flag=ok&catagory=<%=catagory%>"><%end if%>下一页</a></td>
      <td width="17%">第<%=page%>页/共有<%=pgCount%>页</td>
      <td width="17%">第 
        <input type="hidden" name="flag" value="ok">
        <input type=hidden name="catagory" value="<%=catagory%>">
        <select name="page" onchange="frm1.submit();">
          <option value="no"> <%for i=1 to pgCount%> 
          <option value=<%=i%>><%=i%></option>
          <%next%> 
        </select>
        页</td>
    </tr> </table>
 </form>
 <form name=idfrm action="delarticles.asp" method="POST" onsubmit="if(deleted>0){return(confirm('确实要删除这些文章么？'))};if(deletedall>0){return(confirm('确实要删除这个主题么，所有回复将一并删除？'))}">
 <input type=hidden name="catagory" value="<%=catagory%>">    
    <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolordark="#000099" bordercolorlight="#FFFFFF">
    <tr>
      <td> 
        <ul>
          <%if rs.EOF then
         Response.Write "<div align='center'>对不起，还没有任何文章发表。</div>"
      end if   
      do while not rs.EOF and total<pgSz
         total=total+1
         id=rs("id")
         mark=""
         count=count+1
         if trim(rs("content"))="" or isnull(rs("content")) then
            mark="（内空）"
         end if%> 
          <%if cint(rs("fatherid"))<>cint(rs("id")) then%>
              <ul><li>
          <%else%>
            <li>
          <%end if   
              author=trim(rs("author"))
              if not author="" then
                 temp=split(author,"(")
                 username=temp(0)
              end if  %> 
              <a href="detail.asp?id=<%=id%>" onclick="return js_callpages(this.href)"><%=trim(rs("title"))&mark%></a><b>&nbsp; [<a href="javascript:changevalue('<%=username%>');" ><%=author%></a>]</b>&nbsp;
              <small>(点中<font color="red"><%=rs("readtimes")%></font>次，发表时间:<%=trim(rs("ondate"))&trim(rs("ontime"))%>) 操作： 
                 <select name="act_<%=id%>" onChange="if(this.value=='delete'){deleted++;}else{if(deleted>=0){deleted--;}}if(this.value=='deleteall'){deletedall++;}else{if(deletedall>=0){deletedall--;}}">
                       <option></option>
                       <%if rs("score")=1 then%> 
                       <option value="normal">普通</option>
                       <%else%> 
                       <option value="good">精华</option>
                       <%end if%> 
                       <option value="delete">删除</option>
                       <%if cint(rs("fatherid"))=cint(rs("id")) then%>
                       <option value="deleteall">主题删除</option>
                       <%end if%>
                 </select>
              </small>
           <%if cint(rs("fatherid"))<>cint(rs("id")) then%>
              </li></ul>
          <%else%>
            </li>
          <%end if%>
         
          <%rs.MoveNext
      loop
      '把数据库恢复成原样
      sql="update articles set fatherid=0 where id=fatherid"
      conn.Execute sql
      conn.CommitTrans
      %> 
        </ul>
      </td>
  </tr>
</table>
  <p align="center"> 
    <input type="hidden" name="page" value=<%=page%>>
    <input type="hidden" name="count" value=<%=count%>>
    <input type="submit" name="Submit" value="确定">
    <input type="reset" name="Submit2" value="恢复">
  </p>
</form>
<form name="frmpower" action="changepower.asp" method="POST">
  <p align="center">封闭 
    <input type="text" name="username">
    <input type=hidden name="catagory" value="<%=catagory%>">
    在本版发表文章的权限。 
    <input type="submit" name="Submit3" value="确定">
    <input type="reset" name="Submit4" value="恢复">
  </p>
  <p align=center><a href="changepower.asp?catagory=<%=catagory%>">解除封闭</a></p>
</form>
</body>
</html>
