<!--#include file="../../../VAR.asp"-->
<!--#include file="../../function.inc"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>处理结果</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
DIV,p {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
TD {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
A {COLOR: #003366; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #ff0000; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
SELECT {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
input {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
</STYLE>
</head>
<body bgcolor="#FFFFFF">
<%catagory=request("catagory")
  directory=request("dir")
  'Response.Write session("pass"&catagory)&catagory
  if (not session("pass"&catagory)) and not session("admin") or session("administrator")="" then
     Response.Write "<div align=center>对不起，您不能打开本页！</div>"
     Response.End
  end if
  set conn=server.CreateObject("ADODB.Connection")
  Conn.open OpenString
  if request("createdir_confirm")="确定" then
     'Response.Write request("dirname")
     title=SqlTran(request("dirname"))
     if instr(1,title,"\")<>0 or instr(1,title,"'")<>0 then
        Response.Write "<div align=center>对不起，文件夹名称中不能含有“<font color=red>\</font>”或“<font color=red>'</font>”"
        Response.Write "<p><a href='javascript:history.go(-1)'>返回</a></p></div>"
        Response.End
     end if   
     location=directory
     deal_man=session("administrator")
     deal_date=date
     mytype=1
     sql="select * from elite where location='"&directory&"' and title='"&title&"'"
     set rs=conn.Execute(sql)
     if not rs.eof then
        Response.Write "<div align=center>建立文件夹重名，请<a href='javascript:history.go(-1)'>重新再试</a>！</div>"
        Response.End
     end if
     sql="insert into elite (title,catagory,deal_man,type,location,deal_date) values ('"_
        &title&"','"&catagory&"','"&deal_man&"',"&mytype&",'"&location&"','"&deal_date&"')"
     'Response.Write sql
     'Response.End 
     conn.Execute sql   
%>
<div align="center">
  <p>建立文件夹成功！</p>
  <p><a href="<%=session("address")%>">请返回上一页</a></p>
  </div>
<%session("address")=""
end if
if request("mainconfirm")="确定" or request("viewconfirm")="确定" then
   set rs=server.CreateObject("ADODB.Recordset")
   conn.Errors.Clear
   on error resume next
   for each j in Request.Form
      'Response.Write j&"="&request(j)&"<br>"
      if left(j,14)="articleaction_" and request(j)<>"" then
         count=count+1
         t=split(j,"_")
         id=t(1)
         '查找一下这篇文章是否已经处理过了。
         sql="select count(*) as cn from elite where id="&id
         '如果是加入精华区：
         if request("mainconfirm")="确定" then
            sql="select count(*) as cn from elite where article_id="&id
         end if   
         set rstemp=conn.Execute(sql)
         if rstemp(0)>0 then
            if request("mainconfirm")="确定" then
                Response.Write "<div align=center>对不起，这篇文章：<font color=red>"&id&"</font>已经被移动到相应的文件夹了！</div>"
                count=count-1
            else
            '如果在已整理的精华区里面的话
                location=request(j)
                sql="update elite set location='"&location&"' where id="&id
                conn.Execute sql
                %>
				<div align="center">文章：<font color=blue><%=id%></font>移动到文件夹：<font color=red><%=location%></font><br></div>
                <% 
            end if        
         else
            '如果是加入精华区：
           if request("mainconfirm")="确定" then
				sql="select * from articles where id="&id
				rs.Open sql,conn,3 
				if rs.eof then 
				   Response.Write "<div align=center>对不起，这篇文章：<font color=red>"&id&"</font>已经被删除！</div>"
				end if
				title=SqlTran(trim(rs("title")))
				deal_man=session("administrator")
				author=rs("author")
				if author<>"" then
					t=split(author,"(")
					author=t(0)
				end if    
				location=request(j)
				deal_date=date
				rs.close
				sql="insert into elite (title,catagory,deal_man,author,article_id,type,location,deal_date) values('"_
					 &title&"','"&catagory&"','"&deal_man&"','"&author&"',"&id&",0,'"&location&"','"&deal_date&"')"
				conn.Execute(sql)
	       end if
	       if request("viewconfirm")="确定" then
	            location=request(j)
				deal_date=date
				deal_man=session("administrator")
	            sql="update elite set location='"&location&"',deal_date='"&deal_date&"',deal_man='"&deal_man&"' where id="&id
	            conn.Execute sql
	       end if       			
         %>
				<div align="center">文章：<font color=blue><%=id%></font>移动到文件夹：<font color=red><%=location%></font><br>
  <%     end if
         rstemp.close
         set rstemp=nothing  
     end if
      if left(j,11)="delarticle_" and request(j)="del" then
         delcount=delcount+1
         t=split(j,"_")
         id=t(1)
         sql="delete from elite where id="&id
         conn.Execute sql
         %><div align=center>文章：<font color=blue><%=id%></font>被删除。<br>
    <%end if
      if  left(j,7)="rename_" and request(j)<>"" then  
          renamecount=renamecount+1
          t=split(j,"_")
          id=t(1)
          title=SqlTran(request(j))
          deal_man=session("administrator")
          deal_date=date
          sql="update elite set title='"&title&"',deal_man='"&deal_man&"',deal_date='"&deal_date&"' where id="&id
          'Response.Write sql
          conn.Execute sql
          %><div align=center>文章：<font color=blue><%=id%></font>被改名为：<font color=red><%=title%></font>。<br>
      <%end if    
   next
   if conn.Errors.Count>0 then
     for i=0 to conn.Errors.Count
       Response.Write conn.Errors(i).Description&"<br>"
     next
     Response.End
   end if%>
</div>
<%if count>0 then%><p align="center"><font color=red><%=count%></font>篇文章移动成功！</p><%end if%>
<%if delcount>0 then%><p align="center"><font color=red><%=delcount%></font>篇文章删除成功！</p><%end if%>
<%if renamecount>0 then%><p align="center"><font color=red><%=renamecount%></font>篇文章改名成功！</p><%end if%>
<p align="center"><a href="<%=session("address")%>">请返回上一页</a></p>
<%end if%>
</body>
</html>
