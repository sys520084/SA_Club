<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="javascript">
function js_callpage(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=330");
  newwin.focus();
  return false;
}  
</script>
<STYLE type=text/css>
DIV,p {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
TD {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
A {COLOR: #003366; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #ff0000; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
SELECT {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
input {FONT-FAMILY: "����"; FONT-SIZE: 9pt}
</STYLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></HEAD>
<BODY bgcolor="#003366" text="#FFFFFF" link="#CCFF00" vlink="#CCFF00" alink="#CCFF00">
<%sub checkvalid(txt)
     if len(txt)>0 then
        for i=1 to len(txt)
           if mid(txt,i,1)="'" then
              Response.Write "<div align=center>�Բ��𣬿����в��ܰ����ַ���<font color=green>'</font>����<a href='javascript:history.go(-1)'><font color='#ccff00'>������һ��</font></a>��</div></center>"
              Response.End 
           end if
        next 
     end if
  end sub            
  if request("changepass")="ȷ��" then
     name=request("name")
     oldpass=request("oldpass")
     set conn=server.CreateObject("ADODB.Connection")
     conn.Open OpenString
     sql="select username,passwd from users where username='"&name&"'" 
     set rs=conn.Execute(sql)
     if rs.eof then
        Response.Write "<div align=center>�Բ����û����������<a href='javascript:history.go(-1)'><font color='#ccff00'>����������һ��</font></a></div>"
        Response.End
     end if
     if trim(rs("passwd"))<>oldpass then
        Response.Write "<div align=center>�Բ�����������������<a href='javascript:history.go(-1)'><font color='#ccff00'>����������һ��</font></a></div>"
        Response.End
     end if
     checkvalid(request("newpass"))
     if request("newpass")<>"" then
        sql="update users set passwd='"&request("newpass")&"' where username='"&name&"'"
        'Response.Write sql
        'response.end
        conn.Execute(sql)
     end if
   end if  %> 
<P align=center><font color="#FFFFFF">�����޸ĳɹ���</font></P>

</BODY>
</HTML>