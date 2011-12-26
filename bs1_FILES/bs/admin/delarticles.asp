<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
-->
</style>
</HEAD>
<BODY bgcolor="#FFFFFF">
<%catagory=request("catagory")
if (not session("pass"&catagory)) and not session("admin") then
   Response.Write "<div align=center>对不起，您不能打开本页！</div>"
   Response.End
end if
   set conn=server.CreateObject("ADODB.Connection")
    conn.Open OpenString
    set rs1=server.CreateObject("ADODB.Recordset")
    good=0
    delete=0
    normal=0
for each j in Request.Form
  conn.BeginTrans
  conn.Errors.Clear 
  markscore=0
  if left(j,4)="act_" then
          id=right(j,len(j)-4)
    'Response.Write id
    sql="select author from articles where id="&id
    set rs=conn.Execute(sql)
    if not rs.eof then
       author=rs(0)
       'Response.Write author
       if author<>"" then
          if instr(1,author,"(")=0 then
             username=author
          else   
             username=left(author,instr(1,author,"(")-1)
          end if
       end if
       'Response.Write username
       'Response.End 
    end if 
    rs.close
    select case request(j)
    case "normal"   
       normal=normal+1
       sql="update articles set score=0 where id="&id
    case "good"
       good=good+1
       sql="update articles set score=1 where id="&id
    case "delete"   
       delete=delete+1
       sql="select score,fatherid from articles where id="&id
       set rs=conn.Execute(sql)
       if not rs.eof then
           markscore=rs("score")
           'Response.Write markscore&rs("score")&","
           if isnull(markscore) then
              markscore=0
           end if
           fatherid=cint(rs("fatherid"))
       end if
       rs.close
       set rs=conn.Execute("select id from articles where fatherid="&id)
       if not rs.eof then
          idd=rs(0) 
          sql="update articles set fatherid="&idd&" where fatherid="&id
          fatherid=cint(idd)
          conn.Execute sql
          sql="update articles set fatherid=0 where fatherid=id and fatherid="&idd
          'Response.Write sql
          conn.Execute sql
       end if
       rs.close
       if fatherid<>0 then
          sql="select count(*) from articles where fatherid="&fatherid 
          set rs=conn.Execute(sql)
          children=cint(rs(0))-1
          if children<0 then
             children=0
          end if   
          'Response.Write children
          rs.close
          'Response.Write sql
          'Response.End 
          if children>=0 then
             sql="update articles set children="&children&" where id="&fatherid&" or fatherid="&fatherid
             conn.Execute sql
          end if   
       end if         
       sql="delete from articles where id="&id
    case "deleteall"
          delcount=0
          sql="select id,score,fatherid,author from articles where fatherid="&id&" or id="&id
          set rs=server.CreateObject("ADODB.Recordset")
          'Response.Write sql
          rs.Open sql,conn,1,3 
          do while not rs.eof
				 author=rs("author")
				 delallcount=delallcount+1
				 'Response.Write author
				 if author<>"" then
				   if instr(1,author,"(")=0 then
				      username=author
				   else   
				      username=left(author,instr(1,author,"(")-1)
				   end if
				end if  
				'Response.Write cint(rs("id"))&":"&cint(rs("fatherid"))&"ok"
				if rs("score")=1 and cint(rs("id"))<>cint(rs("fatherid")) then
				   rs("fatherid")=0
				else
				   set rstemp=server.CreateObject("ADODB.Recordset")
				   sql="select articles from users where username='"&username&"'"
				   rstemp.Open sql,conn,1,3
				   rstemp("articles")=rstemp("articles")-1
				   rstemp.Update
				   rstemp.Close
				   set rstemp=nothing
				      
				'Response.Write username
				'Response.End 
				end if
				rs.Update  
				rs.MoveNext
		   loop  
    	rs.Close 
	 	set rs=nothing
         sql="delete articles where fatherid="&id&" or id="&id 
         delete=delete+delallcount
         'Response.Write sql
         'Response.end
   end select   
    'Response.Write sql
    'Response.End
    conn.Execute sql
    if not username="" then
       sql="select articles,markscore from users where username='"&username&"'"
       'Response.Write sql
       'Response.End 
       'on error resume next
       rs1.Open sql,conn,1,3 
       if not rs1.EOF then
         select case request(j)
         case "good"
            rs1("markscore")=rs1("markscore")+1
         case "normal"
            rs1("markscore")=rs1("markscore")-1
         case "delete"
			'Response.Write markscore
			rs1("markscore")=rs1("markscore")-markscore
            'Response.Write rs1("articles")
            rs1("articles")=rs1("articles")-1
         end select    
         rs1.Update
       end if  
       rs1.Close
    end if
  end if       
	if conn.Errors.Count>0 then
	   Response.Write "<div align=center>对不起，您的操作有错误，可能是文章没找到，或者是数据库的某些值不能更新。</div>"
	   Response.Write conn.Errors(0).Description 
	   Response.End
	   conn.RollbackTrans
	else      
	   conn.CommitTrans
	end if 
 next    
 '如果删除了某个纪录，就更新所有的lastdate字段
 if delete>0 then
 	'更新数据库中的lastdate字段
    sql1="select id from articles where fatherid=0 and catagory='"&catagory&"'"
	set rstmp=conn.Execute(sql1)
	do while not rstmp.eof
		tmpid=rstmp(0)
		sql1="select ondate,ontime from articles where fatherid=" & tmpid &" order by id desc"
		set rstmp2=conn.Execute(sql1)
		if not rstmp2.eof then
			sql1="update articles set lastdate='" & rstmp2(0) &" " & rstmp2(1) & "' where id=" & tmpid
		else
			sql1="update articles set lastdate=ondate where id=" & tmpid
		end if
		conn.Execute sql1
		rstmp.movenext
	loop
	rstmp.close
 end if
 %> 
<center><P>操作成功，<%=delete%>篇文章被删除。<%=good%>篇文章被标为好，<%=normal%>篇文章被标为一般。</P>
<p><a href="admin.asp?catagory=<%=catagory%>&page=<%=cint(request("page"))%>&flag=ok">返回上一页</a></p>
</center>
</BODY>
</HTML>
