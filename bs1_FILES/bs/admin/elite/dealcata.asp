<!--#include file="../../../VAR.asp"-->
<!--#include file="../../function.inc"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
DIV,p {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
TD {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
A {COLOR: #003366; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #ff0000; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
SELECT {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
input {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
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
     Response.Write "<div align=center>�Բ��������ܴ򿪱�ҳ��</div>"
     Response.End
   end if
   set conn=server.CreateObject("ADODB.Connection")
   conn.Open OpenString
   sql="select * from elite where id="&id
   set rs=conn.Execute(sql)
   if rs.eof then 
      Response.Write "<div align=center>��ȡĿ¼ʧ�ܣ����������Ŀ¼�Ѿ���ɾ����</div>"
      Response.End
   end if  
   directory=trim(rs("location"))
   title=trim(rs("title"))
   set rs=nothing
if not request("dealcataconfirm")="ȷ��" then
%> 
<form method="post" action="dealcata.asp">
  <div align="center">
    <p>
     <% if request("action")="rename" then%> ���ļ�������Ϊ��
      <input type="text" name="directory" value=<%=title%> maxlength="100">
      <%end if
     if request("action")="move" then
        sql="select id,title,location from elite where type=1 and (not location like '"&directory&title&"%') and catagory='"&catagory&"' and not id="&id&" order by location,title"
        'Response.Write sql
        set rs=conn.Execute(sql)
        %></p>
       
		<p>��<font color=red><%=directory&title%></font> �ƶ��������ļ��У� 
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
		<p>�������»��ļ��н�һ�����ߣ�</p>
    <%color="blue"
      end if%>
    <%if request("action")="del" then
         color="red"%>
        <p>ȷʵҪɾ����Щ����ô�� </p>
        <p>�������»��ļ��н���ɾ����</p>
    <%end if    
      if request("action")="move" or request("action")="del" then%>
		 <table width="100%" border="0">
		<%sql="select * from elite where location like '"&directory&title&"\%' and catagory='"&catagory&"' order by type,location"
		  if request("action")="del" then
			%><tr align="center">
				<td> <font color=<%=color%>><%=directory&title%>���ļ��У�</font> </td>
			  </tr><%
		  end if  
		    set rs=conn.Execute(sql)
		    do while not rs.eof
		    %> 
    	  <tr align="center">
		   <td> <font color=<%=color%>><%if rs("type")=1 then%><%=trim(rs("location"))&trim(rs("title"))%>���ļ��У�<%end if%> 
          <%if rs("type")=0 then%><%=trim(rs("title"))%>�����£�<%end if%></font> </td>
		  </tr>
		  <%rs.movenext
		   loop%>
		</table>
		
    <%end if%> 
    <p>
      <input type=hidden name="action" value=<%=request("action")%>>
      <input type=hidden name="catagory" value="<%=catagory%>">
      <input type=hidden name="id" value=<%=id%>>
      <input type="submit" name="dealcataconfirm" value="ȷ��">
      <input type="reset" name="Submit2" value="��д">
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
            Response.Write "<div align=center>�Բ����ļ��������в��ܺ��С�<font color=red>\</font>����<font color=red>'</font>��"
            Response.Write "<p><a href='javascript:history.go(-1)'>����</a></p></div>"
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
  <p>�ļ��д���ɹ�</p>
  <p><a href="javascript:window.close();">�رմ˴���</a></p>
</div>
<%end if%>
</body>
</html>
