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
</head>
<body bgcolor="#FFFFFF">
<%catagory=request("catagory")
  directory=request("dir")
  'Response.Write session("pass"&catagory)&catagory
  if (not session("pass"&catagory)) and not session("admin") or session("administrator")="" then
     Response.Write "<div align=center>�Բ��������ܴ򿪱�ҳ��</div>"
     Response.End
  end if
  set conn=server.CreateObject("ADODB.Connection")
  Conn.open OpenString
  if request("createdir_confirm")="ȷ��" then
     'Response.Write request("dirname")
     title=SqlTran(request("dirname"))
     if instr(1,title,"\")<>0 or instr(1,title,"'")<>0 then
        Response.Write "<div align=center>�Բ����ļ��������в��ܺ��С�<font color=red>\</font>����<font color=red>'</font>��"
        Response.Write "<p><a href='javascript:history.go(-1)'>����</a></p></div>"
        Response.End
     end if   
     location=directory
     deal_man=session("administrator")
     deal_date=date
     mytype=1
     sql="select * from elite where location='"&directory&"' and title='"&title&"'"
     set rs=conn.Execute(sql)
     if not rs.eof then
        Response.Write "<div align=center>�����ļ�����������<a href='javascript:history.go(-1)'>��������</a>��</div>"
        Response.End
     end if
     sql="insert into elite (title,catagory,deal_man,type,location,deal_date) values ('"_
        &title&"','"&catagory&"','"&deal_man&"',"&mytype&",'"&location&"','"&deal_date&"')"
     'Response.Write sql
     'Response.End 
     conn.Execute sql   
%>
<div align="center">
  <p>�����ļ��гɹ���</p>
  <p><a href="<%=session("address")%>">�뷵����һҳ</a></p>
  </div>
<%session("address")=""
end if
if request("mainconfirm")="ȷ��" or request("viewconfirm")="ȷ��" then
   set rs=server.CreateObject("ADODB.Recordset")
   conn.Errors.Clear
   on error resume next
   for each j in Request.Form
      'Response.Write j&"="&request(j)&"<br>"
      if left(j,14)="articleaction_" and request(j)<>"" then
         count=count+1
         t=split(j,"_")
         id=t(1)
         '����һ����ƪ�����Ƿ��Ѿ�������ˡ�
         sql="select count(*) as cn from elite where id="&id
         '����Ǽ��뾫������
         if request("mainconfirm")="ȷ��" then
            sql="select count(*) as cn from elite where article_id="&id
         end if   
         set rstemp=conn.Execute(sql)
         if rstemp(0)>0 then
            if request("mainconfirm")="ȷ��" then
                Response.Write "<div align=center>�Բ�����ƪ���£�<font color=red>"&id&"</font>�Ѿ����ƶ�����Ӧ���ļ����ˣ�</div>"
                count=count-1
            else
            '�����������ľ���������Ļ�
                location=request(j)
                sql="update elite set location='"&location&"' where id="&id
                conn.Execute sql
                %>
				<div align="center">���£�<font color=blue><%=id%></font>�ƶ����ļ��У�<font color=red><%=location%></font><br></div>
                <% 
            end if        
         else
            '����Ǽ��뾫������
           if request("mainconfirm")="ȷ��" then
				sql="select * from articles where id="&id
				rs.Open sql,conn,3 
				if rs.eof then 
				   Response.Write "<div align=center>�Բ�����ƪ���£�<font color=red>"&id&"</font>�Ѿ���ɾ����</div>"
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
	       if request("viewconfirm")="ȷ��" then
	            location=request(j)
				deal_date=date
				deal_man=session("administrator")
	            sql="update elite set location='"&location&"',deal_date='"&deal_date&"',deal_man='"&deal_man&"' where id="&id
	            conn.Execute sql
	       end if       			
         %>
				<div align="center">���£�<font color=blue><%=id%></font>�ƶ����ļ��У�<font color=red><%=location%></font><br>
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
         %><div align=center>���£�<font color=blue><%=id%></font>��ɾ����<br>
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
          %><div align=center>���£�<font color=blue><%=id%></font>������Ϊ��<font color=red><%=title%></font>��<br>
      <%end if    
   next
   if conn.Errors.Count>0 then
     for i=0 to conn.Errors.Count
       Response.Write conn.Errors(i).Description&"<br>"
     next
     Response.End
   end if%>
</div>
<%if count>0 then%><p align="center"><font color=red><%=count%></font>ƪ�����ƶ��ɹ���</p><%end if%>
<%if delcount>0 then%><p align="center"><font color=red><%=delcount%></font>ƪ����ɾ���ɹ���</p><%end if%>
<%if renamecount>0 then%><p align="center"><font color=red><%=renamecount%></font>ƪ���¸����ɹ���</p><%end if%>
<p align="center"><a href="<%=session("address")%>">�뷵����һҳ</a></p>
<%end if%>
</body>
</html>
