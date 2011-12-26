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
<body bgcolor="#FFFFFF">
<% id=request("id")
   catagory=request("catagory")
   if (not session("pass"&catagory)) and not session("admin") or session("administrator")="" then 
     Response.Write "<div align=center>对不起，您不能打开本页！</div>"
     Response.End
   end if
   set conn=server.CreateObject("ADODB.Connection")
   conn.Open OpenString
   sql="select * from elite where id="&id
   set rs=conn.Execute(sql)
   if rs.eof then 
      Response.Write "<div align=center>提取目录失败，可能是这个目录已经被删除！</div>"
      Response.End
   end if  
   directory=trim(rs("location"))
   title=trim(rs("title"))
   set rs=nothing
if not request("dealcataconfirm")="确定" then
%> 
<form method="post" action="dealcata.asp">
  <div align="center">
    <p>
     <% if request("action")="rename" then%> 改文件夹名称为：
      <input type="text" name="directory" value=<%=title%> maxlength="100">
      <%end if
     if request("action")="move" then
        sql="select id,title,location from elite where type=1 and (not location like '"&directory&title&"%') and catagory='"&catagory&"' and not id="&id&" order by location,title"
        'Response.Write sql
        set rs=conn.Execute(sql)
        %></p>
       
		<p>把<font color=red><%=directory&title%></font> 移动到下列文件夹： 
		  <select name="directory">
		    <option value="\">\</option>
		    <%do while not rs.eof 
		             v=trim(rs("location"))&trim(rs("title"))&"\"%> 
		    <option value="<%=v%>"><%=v%></option>
		    <%rs.movenext
		      loop
		  set rs=nothing    %> 
		  </select>
		</p>
		<p>下列文章或文件夹将一起被移走：</p>
    <%color="blue"
      end if%>
    <%if request("action")="del" then
         color="red"%>
        <p>确实要删除这些文章么？ </p>
        <p>下列文章或文件夹将被删除：</p>
    <%end if    
      if request("action")="move" or request("action")="del" then%>
		 <table width="100%" border="0">
		<%sql="select * from elite where location like '"&directory&title&"\%' and catagory='"&catagory&"' order by type,location"
		  if request("action")="del" then
			%><tr align="center">
				<td> <font color=<%=color%>><%=directory&title%>（文件夹）</font> </td>
			  </tr><%
		  end if  
		    set rs=conn.Execute(sql)
		    do while not rs.eof
		    %> 
    	  <tr align="center">
		   <td> <font color=<%=color%>><%if rs("type")=1 then%><%=trim(rs("location"))&trim(rs("title"))%>（文件夹）<%end if%> 
          <%if rs("type")=0 then%><%=trim(rs("title"))%>（文章）<%end if%></font> </td>
		  </tr>
		  <%rs.movenext
		   loop%>
		</table>
		
    <%end if%> 
    <p>
      <input type=hidden name="action" value=<%=request("action")%>>
      <input type=hidden name="catagory" value="<%=catagory%>">
      <input type=hidden name="id" value=<%=id%>>
      <input type="submit" name="dealcataconfirm" value="确定">
      <input type="reset" name="Submit2" value="重写">
    </p>
    </div>
</form>
<%else
     'conn.BeginTrans 
     select case request("action")
     case "move"
         location_d=request("directory")
         dealman=session("administrator")
         dealdate=date
         sql="select * from elite where location like '"&directory&title&"\%' and catagory='"&catagory&"' order by type"
         set rstemp=conn.Execute(sql)
         do while not rstemp.eof 
            move_count=move_count+1
            changed_id=rstemp("id")
            temploc=trim(rstemp("location"))
            if instr(1,temploc,directory)<>0 then
               location_c=location_d&right(temploc,len(temploc)-len(directory)) 
               sql="update elite set location='"&location_c&"',deal_man='"&dealman&"',deal_date='"&dealdate&"' where id="&changed_id
               Response.Write sql
               conn.Execute sql
            end if
            rstemp.movenext
         loop
         sql="update elite set location='"&location_d&"',deal_man='"&dealman&"',deal_date='"&dealdate&"' where id="&id
         conn.Execute sql
     case "rename"
         'Response.Write directory
         if instr(1,request("directory"),"\")<>0 or instr(1,request("directory"),"'")<>0then
            Response.Write "<div align=center>对不起，文件夹名称中不能含有“<font color=red>\</font>”或“<font color=red>'</font>”"
            Response.Write "<p><a href='javascript:history.go(-1)'>返回</a></p></div>"
            Response.End
         end if 
         location_d=directory&request("directory")&"\"
         dealman=session("administrator")
         dealdate=date
         sql="select * from elite where location like '"&directory&title&"\%' and catagory='"&catagory&"' order by type"
         directory=directory&title&"\"
         set rstemp=conn.Execute(sql)
         do while not rstemp.eof 
            move_count=move_count+1
            changed_id=rstemp("id")
            temploc=trim(rstemp("location"))
            if instr(1,temploc,directory)<>0 then
               location_c=location_d&right(temploc,len(temploc)-len(directory)) 
               sql="update elite set location='"&location_c&"',deal_man='"&dealman&"',deal_date='"&dealdate&"' where id="&changed_id
               'Response.Write sql
               conn.Execute sql
            end if
            rstemp.movenext
         loop
         title=SqlTran(request("directory"))
         sql="update elite set title='"&title&"',deal_man='"&dealman&"',deal_date='"&dealdate&"' where id="&id
         conn.Execute sql
     case "del"
         sql="delete from elite where id="&id
         conn.Execute sql
         sql="delete from elite where location like '"&directory&title&"\%' and catagory='"&catagory&"'"
         conn.Execute sql
     end select
     'conn.CommitTrans 
    %>
<p align="center">&nbsp;</p>
<div align="center">
  <p>文件夹处理成功</p>
  <p><a href="javascript:window.close();">关闭此窗口</a></p>
</div>
<%end if%>
</body>
</html>
