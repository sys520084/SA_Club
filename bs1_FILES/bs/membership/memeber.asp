<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<%OpenString=DBBS()
requesttype=request("type")
if request("type")="" or request("type")="new" then
	requesttype="new"
end if
%>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="MSHTML 5.00.2920.0" name=GENERATOR>
<META content=none name="Microsoft Border">
<STYLE type=text/css>
TD {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
}
div {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
}
A {COLOR: #0066FF; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #FF0000; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
input,SELECT {
	FONT-FAMILY: "����";
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
  <%if requesttype="new" then%>
  if (!ValidLength(frm1.username.value,1))
     error(frm1.username,"���������룡");
  else 
  if (!ValidLength(frm1.pass.value,1))
     error(frm1.pass,"���������룡");
  else   
  if (frm1.pass.value!=frm1.pass2.value)
     error(frm1.pass2,"�������������Ч��")
  else
  <%end if%>   
  if (!ValidEmail(frm1.email.value))
     error(frm1.email,"��������ȷ��E-mail��");
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
</HEAD>
<BODY bgColor=#FFFFFF text="#FFFFFF">
<DIV align=center><BR>
  <TABLE align=center border=0 width="95%">
    <TBODY> 
    <TR align="center"> 
      <TD> <%if requesttype="new" then%>
          <b>��һ�������������Ϣ</b> 
          <%else%><b> ���ĸ�����������</b> 
        <%end if%> </TD>
    </TR>
    </TBODY>
  </TABLE>
<HR align=center noShade SIZE=1 width="95%">
<%if request("type")="" or request("type")="new" then
	requesttype="new"
  else
      if session("user")="" then
         Response.Write "<div align='center'>�Բ��������ܸ����û���Ϣ����<a href='relogin.asp'>���µ�½</a>��</div>"
         Response.End
      end if
      name=session("user")
      set conn=server.CreateObject("ADODB.Connection")
      conn.Open OpenString
      name=session("user")
      sql="select nickname from users where username='"&name&"'"
      set rs=conn.execute(sql)
      if not rs.EOF then
         nickname=rs(0)
      end if
      rs.close
      sql="select * from detail where username='"&name&"'"
      set rs=conn.execute(sql)
      if rs.EOF then
         Response.Write "<div align=center>�Բ�������û��Ѿ���ɾ���ˡ�</div>"
         Response.End
      end if   
      question=trim(rs("question"))
      anwser=trim(rs("anwser"))
      truename=trim(rs("name"))
      email=trim(rs("email"))
      homepage=trim(rs("homepage"))
      tel=trim(rs("tel"))
      oicq=trim(rs("oicq"))
	  msn=trim(rs("msn"))
      postno=trim(rs("postno"))
      province=trim(rs("province"))
      sex=rs("sex")
      birthday=trim(rs("birthday"))
      byear=year(birthday)    
      bmonth=month(birthday)
      bday=day(birthday)
      address=trim(rs("address"))
      vocation=trim(rs("vocation"))
      education=trim(rs("education"))
      things=trim(rs("interething"))
      tb=false
   end if
   %>
<form name="form2" id="form2" method="post" action="testuser.asp" target="_blank"><input type="hidden" name="username"></form>

<FORM action="savemember.asp" method="post" name="frm1" id="frm1" onsubmit="return Validata();">
    <div>��<font color="#FF0000">*</font>������д</div>
    <BR>
    <TABLE align=center bgColor=#E2ECF5 border=0 cellPadding=5 cellSpacing=2 
width="86%">
      <TR> 
        <TD align=left vAlign=center width="35%"><font color="#000000"><b>�û�����</b></font></TD>
        <TD width="65%"><font color="#000000"> 
          <%if requesttype="new" then%>
          <input name="username" type="text">
          *&nbsp;&nbsp;&nbsp;&nbsp;
          <input type="button" onclick="form2.username.value=frm1.username.value;form2.submit();" value="�����û����Ƿ����">
          <%else%>
          <input name="username" type="hidden" value="<%=name%>">
          <%=name%></font> <%end if%> </TD>
        <input name="type" type=hidden value="<%=requesttype%>">
      </TR>
      <%if requesttype="new" then%>
      <tr> 
        <td align=left valign=center width="35%"><font color="#000000"><b>���룺</b></font></td>
        <td width="65%"> <font color="#000000"> 
          <input type="password" name="pass" size="20" maxlength="20">
          *</font> </td>
      </tr>
      <TR> 
        <TD align=left vAlign=center width="35%"><font color="#000000"><b>����ȷ�ϣ�</b></font></TD>
        <TD width="65%"> <font color="#000000"> 
          <input type="password" name="pass2" size="20" maxlength="20">
          *</font> </TD>
      </TR>
      <%end if%>
      <TR> 
        <TD align=left vAlign=center><font color="#000000"><STRONG>���������E-mail:</STRONG></font></TD>
        <TD> <font color="#000000"> 
          <input name="email" size=30 maxlength="30" value="<%=email%>">
          *</font> </TD>
      </TR>
      <TR> 
        <TD align=left vAlign=center width="35%"><font color="#000000"><b>�ǳƣ�</b></font></TD>
        <TD width="65%"> <font color="#000000"> 
          <input type="text" name="nickname" size="20" maxlength="20" value=<%=nickname%>>
          </font></TD>
      </TR>
      <TR> 
        <TD align=left vAlign=center width="35%"><font color="#000000"><b>������ʾ���⣺</b></font></TD>
        <TD width="65%"> <font color="#000000"> 
          <input type="text" name="question" size="20" maxlength="20" value=<%=question%>>
          </font></TD>
      </TR>
      <TR> 
        <TD align=left vAlign=center width="35%"><font color="#000000"><b>������ʾ����𰸣�</b></font></TD>
        <TD width="65%"> <font color="#000000"> 
          <input type="text" name="anwser" size="20" maxlength="20" value=<%=anwser%>>
          </font></TD>
      </TR>
      <TBODY>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><STRONG>�����������ʵ����:</STRONG></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <input name="truename" size=20 maxlength="20" value=<%=truename%>>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center 
width="35%"><font color="#000000"><STRONG>��������ĸ�����ҳ:</STRONG></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <input name="homepage" size=30 maxlength="50" value=<%=homepage%>>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><STRONG>���������QQ:</STRONG></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <input name="oicq" maxlength="15" value=<%=oicq%>>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center><font color="#000000"><STRONG>���������MSN:</STRONG></font></TD>
          <TD> <font color="#000000"> 
            <input name="msn" maxlength="15" value=<%=MSN%>>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center width="35%" height="25"> <p><font color="#000000"><b>�绰��</b></font></p></TD>
          <TD height="25" width="65%"> <font color="#000000"> 
            <input type="text" name="tel" maxlength="15" value=<%=tel%>>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><STRONG>��ѡ�������ڵ���: 
            </STRONG></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <SELECT name="province" size=1>
              <%if not province="" then%>
              <option value=<%=province%> selected><%=province%></option>
              <%end if%>
              <option value=""></option>
              <option value="����">����</option>
              <option value="�㶫">�㶫</option>
              <option value="����">����</option>
              <option value="����">����</option>
              <option value="����">����</option>
              <option value="���">���</option>
              <option value="����">����</option>
              <option value="����">����</option>
              <option value="����">����</option>
              <option value="�ӱ�">�ӱ�</option>
              <option value="ɽ��">ɽ��</option>
              <option value="ɽ��">ɽ��</option>
              <option value="������">������</option>
              <option value="����">����</option>
              <option value="�Ϻ�">�Ϻ�</option>
              <option value="����">����</option>
              <option value="�ຣ">�ຣ</option>
              <option value="�½�">�½�</option>
              <option value="����">����</option>
              <option value="����">����</option>
              <option value="�Ĵ�">�Ĵ�</option>
              <option value="����">����</option>
              <option value="����">����</option>
              <option value="���ɹ�">���ɹ�</option>
              <option value="����">����</option>
              <option value="����">����</option>
              <option value="����">����</option>
              <option value="����">����</option>
              <option value="����">����</option>
              <option value="�㽭">�㽭</option>
              <option value="����">����</option>
            </SELECT>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><STRONG>��ѡ������Ա�:</STRONG></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <input <%if sex=0 then%>CHECKED<%end if%> name="sex" type=radio value=0>
            �� 
            <input <%if sex=1 then%>checked<%end if%> name="sex"
      type=radio value=1>
            Ů </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><B>�������ڣ�</B></font></TD>
          <TD width="65%"><font color="#000000"> 
            <input maxLength=4 name="year" size=4 value=<%=byear%>>
            �� 
            <input maxLength=2 
      name="month" size=2 value=<%=bmonth%>>
            �� 
            <input maxLength=2 name="day" size=2 value=<%=bday%>>
            ��</font></TD>
        </TR>
        <tr> 
          <td align=left valign=center width="35%"><font color="#000000"><b>��ϸ��ַ��</b></font></td>
          <td width="65%"> <font color="#000000"> 
            <input type="text" name="address" size="40" maxlength="100" value="<%=address%>">
            </font></td>
        </tr>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><B>����������</B></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <select name="education" size=1>
              <option value="<%=education%>" selected><%=education%></option>
              <option value="��ѧ">��ѧ</option>
              <option value="ר��">ר��</option>
              <option value="����">����</option>
              <option value="˶ʿ">˶ʿ</option>
              <option value="��ʿ">��ʿ</option>
              <option value="����">����</option>
            </select>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><STRONG>���ְҵ:</STRONG></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <select name="vocation" size=1>
              <option value="<%=vocation%>" selected><%=vocation%></option>
              <option value="����ҵ">����ҵ</option>
              <option value="ѧ��">ѧ��</option>
              <option value="���е�λ">���е�λ</option>
              <option value="�����ҵ">�����ҵ</option>
              <option value="����ҵ">����ҵ</option>
              <option value="��ҵ">��ҵ</option>
              <option value="������ҵ">������ҵ</option>
              <option value="����ҵ">����ҵ</option>
              <option value="����ʦ">����ʦ</option>
              <option value="���ܣ�����">���ܣ�����</option>
              <option value="��������">��������</option>
              <option value="����ҵ">����ҵ</option>
              <option value="����/���/�г�">����/���/�г�</option>
              <option value="����">����</option>
            </select>
            </font></TD>
        </TR>
      </TBODY>
    </TABLE>
    <TABLE align=center bgColor=#DBE7F2 border=0 cellPadding=5 cellSpacing=2 
width="86%">
      <TBODY> 
      <TR> 
          <TD align=left vAlign=center width="35%"><STRONG><font color="#000000">����Ȥ��ѧ������ 
            ����ѡ��</font></STRONG></TD>
        </TR>
  <TR> 
          <TD align=middle height="103"> 
		  <%showBack=things
		  if not isnull(showBack) then
		  showBack=replace(showBack,"����ѧ,","")
		  showBack=replace(showBack,"��ѧ,","")
		  showBack=replace(showBack,"����,","")
		  showBack=replace(showBack,"��ѧ,","")
  		  showBack=replace(showBack,"����,","")
 		  showBack=replace(showBack,"�˹�����,","")
		  showBack=replace(showBack,"�˹�����,","")
  		  showBack=replace(showBack,"���ѧ,","")
		  showBack=replace(showBack,"����ѧ,","")
		  showBack=replace(showBack,"����ѧ,","")
		  showBack=replace(showBack,"ϵͳ��ѧ,","")
		  showBack=replace(showBack,"�������ѧ,","")
		  end if
		  %>
            <table width="100%" border="0">
              <tr> 
                <td width="24%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"����ѧ")<>0 then%>
                  checked
                <%end if%> name="hobby" value="����ѧ">
                  ����ѧ</font></td>
                <td width="23%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"��ѧ")<>0 then%>
                  checked
                <%end if%> name="hobby" value="��ѧ">
                  ��ѧ</font></td>
                <td width="26%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"����")<>0 then%>
                  checked
                <%end if%> name="hobby" value="����">
                  ����</font></td>
                <td width="27%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"��ѧ")<>0 then%>
                  checked
                <%end if%> name="hobby" value="��ѧ">
                  ��ѧ</font></td>
              </tr>
              <tr> 
                <td width="24%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"����")<>0 then%>
                  checked
                <%end if%> name="hobby" value="����">
                  ����</font></td>
                <td width="23%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"�˹�����")<>0 then%>
                  checked
                <%end if%> name="hobby" value="�˹�����">
                  �˹�����</font></td>
                <td width="26%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"�˹�����")<>0 then%>
                  checked
                <%end if%> name="hobby" value="�˹�����">
                  �˹�����</font></td>
                <td width="27%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"���ѧ")<>0 then%>
                  checked
                <%end if%> name="hobby" value="���ѧ">
                  ���ѧ </font></td>
              </tr>
              <tr> 
                <td width="24%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"����ѧ")<>0 then%>
                  checked
                <%end if%> name="hobby" value="����ѧ">
                  ����ѧ</font></td>
                <td width="23%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"����ѧ")<>0 then%>
                  checked
                <%end if%> name="hobby" value="����ѧ">
                  ����ѧ</font></td>
                <td width="26%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"ϵͳ��ѧ")<>0 then%>
                  checked
                <%end if%> name="hobby" value="ϵͳ��ѧ">
                  ϵͳ��ѧ </font></td>
                <td width="27%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"�������ѧ")<>0 then%>
                  checked
                <%end if%> name="hobby" value="�������ѧ">
                  �����</font></td>
              </tr>
            </table>
            <p>����:<font color="#000000">
              <input name="otherback" type="text" id="otherback" value="<%=showBack%>" size="40" maxlength="100">
              ����,�Ÿ����� </font></p>
            </TD>
      </TR></TBODY></TABLE><BR>
    <input type="submit" name="complete" value="ȷ��">
    <input type="reset" name="Submit2" value="��д">
  </FORM>
</DIV>
</BODY></HTML>
