<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<!--#include file="../forum/aspFunctions.asp"-->
<%OpenString=DBBS()
detailcontent=""
%>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="MSHTML 5.00.2920.0" name=GENERATOR>
<META content=none name="Microsoft Border">
<STYLE type=text/css>
DIV {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
}
TD {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
}
A {COLOR: #CCFF00; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #CCFF00; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
input,SELECT {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
	background-color: #FFFFFF;
}
</STYLE>
<script language="javascript">
 var errfound;
function ValidLength(item,len){
 return(item.length>=len);
 }

function ValidEmail(item){
    if (item.length==0)
     return false;
    if (item.indexOf('@',0)==-1)
        return false;
    return true;    
 }
function ValidNumber(item){
    if (isNaN(item))return false;
    return ValidLength(item,1);
 }   
function Validata(){
  errfound=false;
  if (!ValidLength(frm1.pass.value,1))
     error(frm1.pass,"请输入密码！");
  else   
  if (frm1.pass.value!=frm1.pass2.value)
     error(frm1.pass2,"你输入的密码无效！")
  else
  if (!ValidEmail(frm1.email.value))
     error(frm1.email,"请输入正确的E-mail！");
   return !errfound;
 }

function error(elem,text){
 if(errfound) return;
 elem.focus();
 window.alert(text);
// elem.select();
 errfound=true;
}

</script>
<title>快速添加新用户</title>
</HEAD>
<BODY text="#FFFFFF">
<DIV align=center><BR>
  <TABLE align=center border=0 width="95%">
    <TBODY> 
    <TR align="center"> 
      <TD> 
	</TD>
    </TR>
    </TBODY>
  </TABLE>
<HR align=center noShade SIZE=1 width="95%">
<%sub checkvalid(txt)
     if len(txt)>0 then
        for i=1 to len(txt)
           if mid(txt,i,1)="'" or mid(txt,i,1)="(" then
              Response.Write "<div align=center>对不起，用户名或口令中不能包含字符：<font color=red>'</font>或<font color=red>(</font>请您<a href='javascript:history.go(-1)'>，请再试一次</a>。</div></center>"
              Response.End 
           end if
        next 
     end if
  end sub            
  if request("type")="new" then
      set conn=server.CreateObject("ADODB.Connection")
      conn.Open OpenString
      'set conn1=server.CreateObject("ADODB.Connection")
      'conn1.Open session("OpenString")
      name=MyhtmlEncode(request("name"))
      pass=MyhtmlEncode(request("pass"))
      checkvalid(name)
      checkvalid(pass)
      'sql="select username from individual where username='"&name&"'"
      'Response.Write sql
      'set rs1=conn1.Execute(sql)
      sql="select username from users where username='"&name&"'"
      set rs=server.CreateObject("ADODB.Recordset")
      rs.Open sql,conn,3
      if not rs.eof then
        Response.Write "<div align='center'>对不起，已经有人使用这个名称了，请您<a href='javascript:history.go(-1)'>再换一个</a>新的。</div>"
        Response.End
      end if
      'if not rs1.eof then
        'Response.Write "<div align='center'>对不起，已经有人使用这个名称了，请您<a href='javascript:history.go(-1)'>再换一个</a>新的。</div>"
        'Response.End
      'end if
      'rs1.close
      'conn1.Close
      'set rs1=nothing
      if name="" or isnull(name) then
        Response.Write "<div align='center'>对不起，用户名不能为空请您<a href='javascript:history.go(-1)'>再试一次</a>。</div>"
        Response.End
      end if
      rs.Close
      conn.BeginTrans
      '开始更新数据库
      rs.Open "users",conn,1,3
      rs.AddNew
      rs("username")=name
      rs("passwd")=pass
      rs("signdate")=date
      rs("lastdate")=date
      rs("times")=0
      rs("markscore")=0
      rs("articles")=0
      rs.Update
      rs.Close
      rs.Open "detail",conn,1,3
      rs.AddNew 
      rs("username")=name
      rs("email")=trim(Request("email"))
      rs("sex")=0
      rs.Update
      rs.Close
      rs.Open "memberdes",conn,1,3
      rs.AddNew 
      rs("username")=name
	  rs("showtitle")=name&"的头脑风暴"
	  rs("detailcontent")=detailcontent
      randomize
      rs("portrait")="pre\"&(int(rnd()*10)+1)&".jpg"
	  

      rs.Update
      rs.Close
      conn.CommitTrans
      session("user")=name
	  UrlTail=MyUrlDecode(request("UrlTail"))
	  %>
      <div align=center>您已经注册成功！随时可以更改您的资料，请<A href="<%=UrlTail%>">继续操作</A>！ </div>
      <script language="javascript">
		//window.location.replace("<%=UrlTail%>");
	  </script>

   <%
   response.end 
   end if
   %>
<FORM name="frm1" action="fastaddnew.asp" method="post" onsubmit="return Validata();">
    <font color="#CCFF00">带</font><font color="#FF0000">*</font><font color="#CCFF00">必须填写</font><BR>
    <TABLE align=center bgColor=#CCCCCC border=0 cellPadding=5 cellSpacing=2 
width="86%">
      <TR> 
        <TD align=left vAlign=center width="35%"><b>用户名：</b></TD>
        <TD width="65%"><input name="name" size="20" maxlength=20></TD>
        <input name="type" type=hidden value="new">
		<input name="UrlTail" type=hidden value="<%=request("UrlTail")%>">
      </TR>
      <tr> 
        <td align=left valign=center width="35%"><b>密码：</b></td>
        <td width="65%"> <input type="password" name="pass" size="20" maxlength="20"> 
          <font color="#FF0000">*</font> </td>
      </tr>
      <TR> 
        <TD align=left vAlign=center width="35%"><b>密码确认：</b></TD>
        <TD width="65%"> <input type="password" name="pass2" size="20" maxlength="20"> 
          <font color="#FF0000">*</font> </TD>
      </TR>
      <TBODY>
        <TR> 
          <TD align=left vAlign=center width="35%"><STRONG>请输入你的E-mail:</STRONG></TD>
          <TD width="65%"> <input name="email" size=30 maxlength="30"> 
            <font color="#FF0000">*</font> </TD>
        </TR>
      </TBODY>
    </TABLE>
    <BR>
    <input type="submit" name="complete" value="确定">
    <input type="reset" name="Submit2" value="重写">
  </FORM>
</DIV>
</BODY></HTML>
