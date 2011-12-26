<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<!--#include file="table.inc"-->
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
<script language=javascript>
function sqlval1(sqls){
   frmquery.sql.value=sqls;
   frmquery.action="buyresult.asp";
   frmquery.target="_blank";
   frmquery.submit();
}  
function sqlval(sqls){
   frmquery.action="queryresult.asp";
   frmquery.sql.value=sqls;
   frmquery.target="_blank";
   frmquery.submit();
}   
</script>
<title>查询结果</title></HEAD>
<BODY bgcolor="#FFFFFF">
<%function EncodeP(Str)
     ReStr=""
     if len(Str)>0 then
        for i=1 to len(Str)
            select case mid(Str,i,1)
            case "'"
               ReStr=ReStr&"''"
            case else
               ReStr=ReStr&mid(Str,i,1)
            end select
        next
     end if
     EncodeP=ReStr
  end function
  if not session("admin") then
     Response.Write "<div align=center>对不起，您不能访问本页！</div>"
     Response.End
  end if              
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  if Request.ServerVariables("REQUEST_METHOD")="POST" then
    dim SubSql()
    temp=0
    SqlPass=false
    for each j in Request.Form
      'Response.Write trim(request(j))
      passthis=false
      if trim(request(j))="" then
        ' Response.Write "ok"
         passthis=true   
      end if 
      if not passthis and not isnull(request(j)) and not left(j,6)="button" then
       'Response.Write  j&":"
       'Response.Write request(j)&"<br>"
           tbProcess=false
           SubSuccess=false
           if left(j,4)="hide" then
             tbProcess=true
             RightPart=right(j,len(j)-4)
             select case RightPart
             case "table"
             table=request(j)
             end select  
           end if
           if j="sql" then
              SubTemp=request(j)
              SqlPass=true
              exit for
           end if    
           if left(j,4)="time" then
              tbProcess=true
              if left(j,8)="timehide" then
                Vname=mid(j,9,len(j)-9)
                'Response.Write Vname+"="
                temp=temp+1
                select case right(j,1)
                case "1" 
                  V1=request(j)
                  'Response.Write V1+"<br>"
                case "2"
                  V2=request(j)
                  'Response.Write V2+"<br>"
                end select
                if temp=2 then
                   SubSuccess=true
                   SubTemp=" ("&Vname&" between '"&V1&"' and '"&V2&"')"
                   temp=0
                end if   
               end if  
            end if
            if left(j,4)="like" then
               tbProcess=true
               SubSuccess=true
               SubTemp=" ("&right(j,len(j)-4)&" like '%"&EncodeP(request(j))&"%')"
            end if   
            if left(j,5)="radio" then
               tbProcess=true
               SubSuccess=true
               SubTemp=" "&right(j,len(j)-5)&"="&request(j)
            end if   
            checkCount=0
            if left(j,5)="check" then
               'Response.Write "ok"
               tbProcess=true
               checkCount=Request.Form(j).Count
               'Response.Write checkCount
               for temp=1 to checkCount
                  Vname=right(j,len(j)-5)
                  V=Request.form(j)(temp)
                  'Response.Write temp
                  if temp=1 then
                     SubTemp=" ("
                  end if   
                  if temp<checkCount then
                     SubTemp=SubTemp&Vname&" like '%"&V&"%' or "
                  end if
                  if temp=checkCount then
                     SubTemp=SubTemp&Vname&" like '%"&V&"%')"
                     SubSuccess=true
                  end if
                next        
            end if
            if not tbProcess then
               Vname=j
               V=EncodeP(request(j))
               tbProcess=true
               SubSuccess=true
               SubTemp=" "&Vname&"='"&V&"'"
            end if
      end if    
      if SubSuccess then
          SubCount=SubCount+1
          redim preserve SubSql(SubCount)
          SubSql(SubCount)=SubTemp
          SubTemp=""
          SubSuccess=false
      end if                  
    next
    if not isnull(SubSql) and subcount>0 then
      for i=1 to ubound(SubSql)
         if i<ubound(SubSql) then
          whereSub=whereSub&subsql(i)&" and"
         else
           whereSub=whereSub&SubSql(i) 
         end if 
          'Response.Write subsql(i)&"<br>"
      next
    end if
    if whereSub="" then
       sql="select * from "&table
    else   
       sql="select * from "&table&" where "&whereSub
    end if   
    if SqlPass then
      sql=SubTemp
      s1=instr(1,sql," from ")
      s2=instr(s1,sql," where ")
        if s2=0 then
            s2=len(sql)
        end if   
         table=mid(sql,s1+5,s2-s1-5)
         table=trim(table)
         'Response.Write table
    else
      session("sql")=sql
    end if
    page=1
end if
    if Request.ServerVariables("REQUEST_METHOD")="GET" then
         page=cint(Request.QueryString("page"))
         sql=session("sql")
         s1=instr(1,sql," from ")
         s2=instr(s1,sql," where ")
         if s2=0 then
            s2=len(sql)
         end if   
         table=mid(sql,s1+5,s2-s1-4)
         table=trim(table)
         if right(table,8)="relation" then
            table="relation"
         end if   
    end if
    PgSz=50
    set rs=server.CreateObject("ADODB.Recordset")
   'Response.Write sql
   'Response.End
    conn.Errors.Clear 
   on error resume next
    rs.Open sql,conn,3
    if conn.Errors.Count>0 then
       for i=0 to conn.Errors.Count-1
          Response.Write conn.Errors(i).Description&"<br>"
       next
          Response.Write "下列语句有错："&sql
          Response.End
    end if      
    if rs.EOF then
       Response.Write "<div align=center>对不起，没有找到符合条件的纪录，<a href='javascript:history.go(-1)'>请重来。</a></div> "
       Response.End
    end if
    rs.PageSize=PgSz
    if page="" then
       page=1
    end if
    rs.AbsolutePage=page
%>
<div align=center>
<div>
  <form name="form1" method="GET" action="queryresult.asp" >
    <%if page>1 then%>&nbsp;&nbsp;<a href="queryresult.asp?page=<%=page-1%>">上一页</a><%end if%>&nbsp;<%if page<rs.pagecount then%><A href="queryresult.asp?page=<%=page+1%>">下一页</a><%end if%>&nbsp;共有<font color="#FF0000"><%=rs.RecordCount%></font>条纪录，察看 
    <select name="page" onchange="form1.submit();">
    <%for i=1 to rs.PageCount%>
      <option <%if page=i then%>selected <%end if%>value=<%=i%>><%=i%></option>
    <%next%>    
    </select>
    页/共<font color="#FF0000"><%=rs.pagecount%></font>页 
  </form>
</div>
<form action="queryresult.asp" name=frmquery method="POST"  >
<input type=hidden name=sql>
<table border="1" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" bordercolorlight="#000033" bordercolordark="#CCFFFF" <%if table="member" or table="detail" then%> width=1000 <%end if%>>
  <tr bgcolor="#333333" align="center"> 
  <%for i=0 to rs.Fields.Count-1%> 
    <td><font color="#CCCCCC"><%=Translate(trim(rs.Fields(i).Name))%></font></td>
  <%next
   if table="detail" then%>
     <td><font color="#CCCCCC">操作</font></td>
 <%end if%>     
  </tr>
<%do while not rs.EOF and ti<PgSz
       ti=ti+1%>
  <tr bgcolor="#FFFFFF"> 
  <%for i=0 to rs.Fields.Count-1%>
    <td><%=ValueTranslate(trim(rs.Fields(i)),table,trim(rs.Fields(i).name))%></td>
  <%next
  if table="detail" then%>
     <td><a href="deluser.asp?username=<%=rs("username")%>" onclick="return confirm('确实要删除这个用户吗？');" target="_Blank">删除</a></td> 
 <%end if
   rs.MoveNext
  loop
%>  
</table>
</form>
<div>
  <form name="form2" method="GET" action="queryresult.asp" >
    <%if page>1 then%>&nbsp;<a href="queryresult.asp?page=<%=page-1%>">上一页</a><%end if%>&nbsp;<%if page<rs.pagecount then%><A href="queryresult.asp?page=<%=page+1%>">下一页</a><%end if%>&nbsp;共有<font color="#FF0000"><%=rs.RecordCount%></font>条纪录，察看 
    <select name="page" onchange="form2.submit();">
    <%for i=1 to rs.PageCount%>
      <option <%if page=i then%>selected <%end if%>value=<%=i%>><%=i%></option>
    <%next%>    
    </select>
    页/共<font color="#FF0000"><%=rs.pagecount%></font>页 
  </form>
</div>
<%rs.Close
  set rs=nothing%>
</div>
</BODY>
</HTML>
