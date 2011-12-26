<%@ Language=VBScript %>
<!-- #include file="../../var.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>版主管理</title>
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:宋体;color:black;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0000CC}
a:visited {text-decoration: none;; color: #660066}
a:hover {text-decoration: underline; color: #FF0000}
a:active {text-decoration;background-color: #CCCCCC; color: #006666}
-->
</style></HEAD>
<BODY bgcolor="#FFFFFF">
<%if not session("admin") then
     Response.Write "<div algin='center'>对不起，您不能打开本页。<a href='jingwen.asp'>请重新登陆</a>"
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
     <p>成功的删除了<%=request("name")%>的<%=request("catagory")%>版的版主之职。</p>
     <p><a href="javascript:history.go(-1);">返回上一页</a>
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
          Response.Write "<div align='center'>对不起，不存在用户："&name&"。您不能指定他为版主。<a href='javascript:history.go(-1)'>请重新指定</a>"
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
         Response.Write "<div algin='center'>对不起，他已经是该版的版主了。</div>" 
         Response.End 
       end if
       rs.close
       set rs=nothing   
  if name<>"" then%> </p>
</div>
<P align="center">成功的封：<%=name%>为：<%=catagory%>版的斑竹。</P>
<%end if
end if%>
<p align=center>所有斑竹列表如下，（点击下列版主名称将会撤销他的版主之职）。</p>
<%sql="select * from power where power=0"
  set rs=conn.Execute(sql)
%><table width="100%" border="0">
  <tr align="center" bgcolor="#999999"> 
    <td>版主名</td>
    <td>管理密码</td>
    <td>负责版</td>
    <td>就任日期</td>
  </tr>
<%do while not rs.eof%> 
  <tr align="center" bgcolor="#CCCCCC"> 
    <td><a href="banzhu.asp?type=del&id=<%=rs("id")%>&name=<%=trim(rs("username"))%>&catagory=<%=trim(rs("catagory"))%>" onclick="return confirm('是否真的删除<%=trim(rs("username"))%>的版主之职？');"> <%=trim(rs("username"))%></a></td>
    <td><%=trim(rs("passwd"))%></td>
    <td><%=trim(rs("catagory"))%></td>
    <td><%=trim(rs("fromdate"))%></td>
  </tr>
<%rs.movenext
  loop%> 
</table>
</BODY>
</HTML>
