<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<!--#include file="aspFunctions.asp"-->
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
<%
UrlTail=MyUrlDecode(request("UrlTail"))
if session("user")="" then
   Response.Write "<div align=center>对不起，您不能打开本页！请<a href='javascript:history.go(-1)'>退回</a></div>"
   Response.End
end if
   set conn=server.CreateObject("ADODB.Connection")
    conn.Open OpenString
    set rs1=server.CreateObject("ADODB.Recordset")
    'conn.BeginTrans
    conn.Errors.Clear 
    id=request("id")
    sql="select * from articles where id="&id
    set rs=conn.Execute(sql)
    if rs.eof then
       Response.Write "<div align=center>对不起，提取文章失败，这篇文章可能已被删除</div>"
       Response.End
    end if
  author=left(rs("author"),len(session("user")))
  username=author
  isadmin=IsSuperAdmin(Conn,Session("user"))
   'Response.Write catagory
  if author<>session("user") and not isAdmin then
     Response.Write "<div align=center>对不起，您不能修改或删除这篇文章</div>"
     Response.End
  end if    
    select case trim(request("action"))
    case "del"   
       markscore=rs("score")
       fatherid=clng(rs("fatherid"))
       rs.close
       set rstemp=conn.Execute("select id from articles where fatherid="&id)
       if not rstemp.eof then
          idd=rstemp(0) 
          sql="update articles set fatherid="&idd&" where fatherid="&id
          fatherid=clng(idd)
          conn.Execute sql
          sql="update articles set fatherid=0 where fatherid=id and fatherid="&idd
          'Response.Write sql
          conn.Execute sql
       end if
       rstemp.close
       if fatherid<>0 then
          sql="select count(*) from articles where fatherid="&fatherid 
          set rstemp=conn.Execute(sql)
          children=cint(rstemp(0))-1
          if children<0 then
             children=0
          end if   
          'Response.Write children
          rstemp.close
          
          'Response.Write sql
          'Response.End 
          if chidren>=0 then
             sql="update articles set children="&children&" where id="&fatherid&" or fatherid="&fatherid
             conn.Execute sql
          end if   
          'Response.Write "ok"
       end if         
       sql="delete from articles where id="&id
    case "edit"
        rs.close
        title=SqlTran(server.HTMLEncode(request("title")))
        content=SqlTran(request("content"))
        face=request("pics")
		catagory=request("catagory")
		tags=SqlTran(request("tags"))
        ondate=date
        ontime=time()
        sql="update articles set title='"&title&"',content='"&content&"',face='"&face&"',lastdate='"&ondate&" "&ontime&"',catagory='"&catagory&"',tags='"&tags&"' where id="&id 
       ' Response.Write request("action")      
    end select   
    'Response.Write sql
    'Response.End
    on error resume next
    conn.Execute sql
    if not username="" then
       sql="select articles,markscore from users where username='"&username&"'"
       'Response.Write sql
       'Response.End 
       'on error resume next
       rs1.Open sql,conn,1,3 
       if not rs1.EOF then
         select case trim(request("action"))
         case "del"
            'Response.Write markscore
            if markscore>0 then
               rs1("markscore")=rs1("markscore")-1
            end if
           ' Response.Write rs1("articles")
            rs1("articles")=rs1("articles")-1
         end select    
         rs1.Update
       end if  
       rs1.Close
    end if
    for i=0 to conn.Errors.Count-1
	   'Response.Write "<div align=center>对不起，您的操作有错误，可能是文章没找到，或者是数据库的某些值不能更新。</div>"
	   Response.Write conn.Errors(i).Description 
	   'Response.End
	next
	if request("action")="del" then
		'更新数据库中的lastdate字段
		'Response.Write "ok"
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
		sql1="delete * from article_category where article_id="&id
		conn.Execute sql1
	end if
		
	'if conn.Errors.Count>0 then   
	'   conn.RollbackTrans
	'else      
	'   conn.CommitTrans
	'end if
 %> 
<center><P>操作成功！</P>
<P>请刷新页面，以查看更改结果</P>
</center>
</BODY>
</HTML>
