<%@ Language=VBScript %>
<!-- #include file="../../var.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>��������</title>
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:����;color:black;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0000CC}
a:visited {text-decoration: none;; color: #660066}
a:hover {text-decoration: underline; color: #FF0000}
a:active {text-decoration;background-color: #CCCCCC; color: #006666}
-->
</style></HEAD>
<BODY bgcolor="#FFFFFF">
<%if not session("admin") then
     Response.Write "<div algin='center'>�Բ��������ܴ򿪱�ҳ��<a href='jingwen.asp'>�����µ�½</a>"
     response.end
  end if
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
 ' Response.Write request("type")&request("id")
  if trim(request("type"))="del" then
  '  Response.Write "ok"
     sql="delete from power where id="&request("id")
     conn.Execute sql
     %> 
     <div align="center">
     <p>�ɹ���ɾ����<%=request("name")%>��<%=request("catagory")%>��İ���ְ֮��</p>
     <p><a href="javascript:history.go(-1);">������һҳ</a>
  <%Response.End 
  end if     
  name=request("name")
  pass=request("pass")
  catagory=request("catagory")
 if name<>"" and catagory<>"" then 
    'Response.Write pass&"ok"
    if isnull(pass) or pass="" then
      ' Response.Write "ok"
       sql="select passwd from users where username='"&name&"'"
       set rs=conn.Execute(sql)
       if rs.eof then
          Response.Write "<div align='center'>�Բ��𣬲������û���"&name&"��������ָ����Ϊ������<a href='javascript:history.go(-1)'>������ָ��</a>"
          response.end
       end if
       pass=trim(rs("passwd"))
       rs.close
       set rs=nothing
    end if    
       sql="select count(*) from power where username='"&name&"' and catagory='"&catagory&"' and power=0"
       set rs=conn.Execute(sql)
       if rs(0)=0 then 
         sql="insert into power (username,passwd,catagory,power,fromdate) values ('"&name&"','"&pass&"','"&catagory&"',0,'"&date&"')"
       '  Response.Write sql
         conn.Execute(sql)
       else
         Response.Write "<div algin='center'>�Բ������Ѿ��Ǹð�İ����ˡ�</div>" 
         Response.End 
       end if
       rs.close
       set rs=nothing   
  if name<>"" then%> </p>
</div>
<P align="center">�ɹ��ķ⣺<%=name%>Ϊ��<%=catagory%>��İ���</P>
<%end if
end if%>
<p align=center>���а����б����£���������а������ƽ��᳷�����İ���ְ֮����</p>
<%sql="select * from power where power=0"
  set rs=conn.Execute(sql)
%><table width="100%" border="0">
  <tr align="center" bgcolor="#999999"> 
    <td>������</td>
    <td>��������</td>
    <td>�����</td>
    <td>��������</td>
  </tr>
<%do while not rs.eof%> 
  <tr align="center" bgcolor="#CCCCCC"> 
    <td><a href="banzhu.asp?type=del&id=<%=rs("id")%>&name=<%=trim(rs("username"))%>&catagory=<%=trim(rs("catagory"))%>" onclick="return confirm('�Ƿ����ɾ��<%=trim(rs("username"))%>�İ���ְ֮��');"> <%=trim(rs("username"))%></a></td>
    <td><%=trim(rs("passwd"))%></td>
    <td><%=trim(rs("catagory"))%></td>
    <td><%=trim(rs("fromdate"))%></td>
  </tr>
<%rs.movenext
  loop%> 
</table>
</BODY>
</HTML>
