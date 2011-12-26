<%@ Language=VBScript %>
<!--#include file="../VAR.asp"-->
<%OpenString=DBBS()
%>

<%
sub checkvalid(txt)
     if len(txt)>0 then
        for i=1 to len(txt)
           if mid(txt,i,1)="'" then
              Response.Write "<div align=center>对不起，用户名或口令中不能包含字符：<font color=red>'</font>请您<a href='javascript:history.go(-1)'>，请再试一次</a>。</div></center>"
              Response.End 
           end if
        next 
     end if
end sub
  session("user")=""
  name=request("name")
  pass=request("pass")
  'response.write name
  'response.write pass
  'response.end
 if name<>"" and pass<>"" then
  checkvalid(name)
  checkvalid(pass)
  name=server.HTMLEncode(name)
  pass=server.HTMLEncode(pass)
  set conn=server.CreateObject("ADODB.Connection")
  conn.Errors.Clear
  on error resume next
  'Response.Write OpenString
  conn.Open OpenString
  'for i=0 to conn.Errors.Count-1
  '    Response.Write conn.Errors(i).Description
  'next  
  'Response.End 

  sql="select * from users where username='"&name&"'"
  set rs=server.CreateObject("ADODB.Recordset")
  'response.write sql
  'response.end
  rs.open sql,conn,3
  'response.end
  jump=false
  if not rs.eof then
     if trim(rs("passwd"))=trim(pass) then
        session("user")=name
        sql="update users set times=times+1, lastdate='"&date&"' where username='"&name&"'"
        'Response.Write sql
		'Response.End 
        'on error resume next
        conn.Execute(sql)

		on error goto 0
		sql="select * from online where username='"&name&"'"
		set rs1=server.CreateObject("ADODB.Recordset")
		'response.write sql
		'response.end
		rs1.open sql,conn,1,3
		'response.end
		if rs1.eof then
			sql="insert into online (username,onlinetime) values ('"&name&"','"&now()&"')"
			'response.write sql
			'response.end
			conn.execute(sql)
		else
				  
			rs1("onlinetime")=now()
			rs1("times")=rs1("times")+1
			rs1.update
		end if
		rs1.close
		set rs1=nothing
        jump=true
     end if
  end if 
  
  rs.Close
  conn.Close
 end if       
if jump then
   Response.Redirect("forum.asp?user="&name) 
end if
%>
<HTML>
<HEAD>
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:宋体;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
-->
</style>
<title>验证用户身份</title></HEAD>
<BODY bgcolor="#CCCC99">
<P align="center">对不起， 您输入的用户名密码不正确。</P>

</BODY>
</HTML>
</P>

</BODY>
</HTML>
