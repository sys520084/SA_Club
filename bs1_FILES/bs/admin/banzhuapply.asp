<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:����;color:black;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0000CC}
a:visited {text-decoration: none;; color: #660066}
a:hover {text-decoration: underline; color: #FF0000}
a:active {text-decoration;background-color: #CCCCCC; color: #006666}
-->
</style>
<script language="javascript">
<!--
function js_callpages(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=300");
  newwin.focus();
  return false;
}
function changevalue(txt){
  frmpower.username.value=txt;
}  
-->
</script></HEAD>
<BODY bgcolor="#FFFFFF">
<%if not session("admin") then
     Response.Write "<div algin='center'>�Բ��������ܴ򿪱�ҳ��<a href='jingwen.asp'>�����µ�½</a>"
     response.end
  end if%>
<table width="100%" border="1" bordercolordark="#FFFFFF" bordercolorlight="#333333" cellpadding="0" cellspacing="0">
  <tr align="center" bgcolor="#CCCCCC"> 
    <td width="9%">�û���</td>
    <td width="8%">��̳</td>
    <td width="47%">ԭ��</td>
    <td width="17%">����ʱ��</td>
    <td width="19%">����</td>
  </tr>
  <%set conn=server.CreateObject("ADODB.Connection")
    conn.Open OpenString
    set rs=server.CreateObject("ADODB.Recordset")
    sql="select * from [application] order by id desc"
    rs.Open sql,conn
    do while not rs.EOF
       j=j+1%> 
    <form name="frm<%=j%>" action="banzhu.asp" method="POST">
 <tr>
      <td width="9%"><a href="../membership/display.asp?name=<%=trim(rs("name"))%>"><%=trim(rs("name"))%></a> 
        <input type="hidden" name="name" value="<%=trim(rs("name"))%>"></td>
      <td width="8%"><%=trim(rs("catagory"))%> 
        <input type="hidden" name="catagory" value="<%=trim(rs("catagory"))%>"></td>
      <td width="47%"><%=trim(rs("reason"))%></td>
      <td width="17%"><%=trim(rs("signdate"))%></td>
      <td width="19%"> 
        <input type=submit name="submit" value="ͨ��">
    <input type=button name="submit" value="ɾ��" onclick="if(confirm('�Ƿ�ȷ��ɾ���˼�¼��')){js_callpages('delapply.asp?id=<%=rs("id")%>')}"></td>
  </tr>
    </form>
<%    rs.movenext
    loop%>
</table><P>&nbsp;</P>
</BODY>
</HTML>
