<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<!--#include file="aspFunctions.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">

<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:宋体;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
-->
</style>
</HEAD>
<BODY bgProperties=fixed bgcolor="#FBFBF4">
<%
   if Request.ServerVariables("REQUEST_METHOD")<>"POST" then
      Response.Write "<div align=center对不起，您不能用这种方式浏览本页。</div>"
      Response.End
   end if   
   if request("save")="确定并提交" then
     name=request("name")
     title=trim(request("title"))
     'Response.Write title
     title=MyhtmlEncode(title)
	 ThisTail=request("ThisTail")
	 UrlTail=MyUrlDecode(request("UrlTail"))
     'Response.Write title
     'response.write "ok;;;;;;"
     if name="" or isnull(name)  then
        Response.Write "<div align='center'>对不起，您不能发表文章</div>"
        Response.end
     end if
     if title="" or isnull(title) then
        Response.Write "<div align='center'>文章标题不能为空！</div>"
        Response.End
     end if
     if len(title)>255 then
        Response.Write "<div align='center'>文章标题过长！</div>"
        Response.End
     end if  
     set conn=server.CreateObject("ADODB.Connection")
     Conn.open OpenString
     sql="select nickname,articles from users where username='"&name&"'"
     set rs1=conn.Execute(sql)
     if rs1.eof then
        Response.Write "<div align='center'>对不起，您不能发表文章</div>"
        Response.End
     end if
     'Response.Write rs1(0)
     nickname=trim(rs1("nickname"))
     paper=cint(rs1("articles"))+1
     if isnull(paper) then
        paper=1
     end if
     sql="update users set articles="&paper&" where username='"&name&"'"
     rs1.close
     conn.BeginTrans
    'response.end
    'on error resume next	
     conn.Execute(sql)
     content=request("content")
     'response.write request("fatherid")
     fatherid=request("fatherid")
	 if fatherid="" then
		fatherid=0
	else
		fatherid=clng(fatherid)
	end if
     ' Response.Write fatherid
     if fatherid<>0 then
        sql="select id from articles where id="&fatherid
        'Response.Write sql
        set rs=conn.Execute(sql)
        if not rs.eof then
          tb=true
        end if
        rs.close
        set rs=nothing
     end if
     if tb then
       sql="update articles set children=children+1,lastdate='"&date&" "&time&"' where id="&fatherid 
       conn.Execute(sql)
     end if  
    set rs=server.CreateObject("ADODB.Recordset")
    rs.Open "articles",conn,1,3
    rs.AddNew
    rs("title")=title
    if nickname="" then
       rs("author")=name
    else
       'Response.Write name&"("&nickname&")" 
       rs("author")=name
    end if
    rs("readtimes")=0
    rs("face")=request("pics")
    rs("capability")=len(content)
    rs("ondate")=date
    rs("ontime")=time
    rs("lastdate")=date&" "&time
    rs("children")=0
    rs("fatherid")=fatherid
    rs("content")=content
	if request("atype")="blog" then
		rs("score")=1
	end if

    catagory=trim(request("catagory"))
	rs("catagory")=catagory
	tags=trim(request("tags"))
	rs("tags")=SqlTran(tags)
    'Response.Write catagory
    conn.errors.clear    
    on error resume next
    rs.Update
    if conn.errors.count>0 then
       for i=0 to conn.errors.count-1
           response.write conn.errors(i).Description
       next
       response.end
     end if
    conn.CommitTrans

%>
<P align="center">发表文章成功，如果想看到您写的文章关闭本窗口后按上面的刷新显示</P>
<P align="center"> <a href="viewtitle.asp?<%=UrlTail%>">返回标题页</a></P>
<script language="JavaScript">
<%if request("IsReply")=1 then%>
	window.close();
<%else%>
    window.location.replace("viewtitle.asp?<%=MyUrlDecode(UrlTail)%>");
<%end if%>
    //document.write(window.opener.location);
</script>
<%end if%>
</BODY>
</HTML>
