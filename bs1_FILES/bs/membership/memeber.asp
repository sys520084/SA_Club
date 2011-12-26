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
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
}
div {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
}
A {COLOR: #0066FF; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #FF0000; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
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
  <%if requesttype="new" then%>
  if (!ValidLength(frm1.username.value,1))
     error(frm1.username,"请输入密码！");
  else 
  if (!ValidLength(frm1.pass.value,1))
     error(frm1.pass,"请输入密码！");
  else   
  if (frm1.pass.value!=frm1.pass2.value)
     error(frm1.pass2,"你输入的密码无效！")
  else
  <%end if%>   
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
</HEAD>
<BODY bgColor=#FFFFFF text="#FFFFFF">
<DIV align=center><BR>
  <TABLE align=center border=0 width="95%">
    <TBODY> 
    <TR align="center"> 
      <TD> <%if requesttype="new" then%>
          <b>第一步，添入个人信息</b> 
          <%else%><b> 您的个人资料设置</b> 
        <%end if%> </TD>
    </TR>
    </TBODY>
  </TABLE>
<HR align=center noShade SIZE=1 width="95%">
<%if request("type")="" or request("type")="new" then
	requesttype="new"
  else
      if session("user")="" then
         Response.Write "<div align='center'>对不起您不能更改用户信息，请<a href='relogin.asp'>重新登陆</a>。</div>"
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
         Response.Write "<div align=center>对不起，这个用户已经被删除了。</div>"
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
    <div>带<font color="#FF0000">*</font>必须填写</div>
    <BR>
    <TABLE align=center bgColor=#E2ECF5 border=0 cellPadding=5 cellSpacing=2 
width="86%">
      <TR> 
        <TD align=left vAlign=center width="35%"><font color="#000000"><b>用户名：</b></font></TD>
        <TD width="65%"><font color="#000000"> 
          <%if requesttype="new" then%>
          <input name="username" type="text">
          *&nbsp;&nbsp;&nbsp;&nbsp;
          <input type="button" onclick="form2.username.value=frm1.username.value;form2.submit();" value="检测该用户明是否可用">
          <%else%>
          <input name="username" type="hidden" value="<%=name%>">
          <%=name%></font> <%end if%> </TD>
        <input name="type" type=hidden value="<%=requesttype%>">
      </TR>
      <%if requesttype="new" then%>
      <tr> 
        <td align=left valign=center width="35%"><font color="#000000"><b>密码：</b></font></td>
        <td width="65%"> <font color="#000000"> 
          <input type="password" name="pass" size="20" maxlength="20">
          *</font> </td>
      </tr>
      <TR> 
        <TD align=left vAlign=center width="35%"><font color="#000000"><b>密码确认：</b></font></TD>
        <TD width="65%"> <font color="#000000"> 
          <input type="password" name="pass2" size="20" maxlength="20">
          *</font> </TD>
      </TR>
      <%end if%>
      <TR> 
        <TD align=left vAlign=center><font color="#000000"><STRONG>请输入你的E-mail:</STRONG></font></TD>
        <TD> <font color="#000000"> 
          <input name="email" size=30 maxlength="30" value="<%=email%>">
          *</font> </TD>
      </TR>
      <TR> 
        <TD align=left vAlign=center width="35%"><font color="#000000"><b>昵称：</b></font></TD>
        <TD width="65%"> <font color="#000000"> 
          <input type="text" name="nickname" size="20" maxlength="20" value=<%=nickname%>>
          </font></TD>
      </TR>
      <TR> 
        <TD align=left vAlign=center width="35%"><font color="#000000"><b>密码提示问题：</b></font></TD>
        <TD width="65%"> <font color="#000000"> 
          <input type="text" name="question" size="20" maxlength="20" value=<%=question%>>
          </font></TD>
      </TR>
      <TR> 
        <TD align=left vAlign=center width="35%"><font color="#000000"><b>密码提示问题答案：</b></font></TD>
        <TD width="65%"> <font color="#000000"> 
          <input type="text" name="anwser" size="20" maxlength="20" value=<%=anwser%>>
          </font></TD>
      </TR>
      <TBODY>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><STRONG>请输入你的真实姓名:</STRONG></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <input name="truename" size=20 maxlength="20" value=<%=truename%>>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center 
width="35%"><font color="#000000"><STRONG>请输入你的个人主页:</STRONG></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <input name="homepage" size=30 maxlength="50" value=<%=homepage%>>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><STRONG>请输入你的QQ:</STRONG></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <input name="oicq" maxlength="15" value=<%=oicq%>>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center><font color="#000000"><STRONG>请输入你的MSN:</STRONG></font></TD>
          <TD> <font color="#000000"> 
            <input name="msn" maxlength="15" value=<%=MSN%>>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center width="35%" height="25"> <p><font color="#000000"><b>电话：</b></font></p></TD>
          <TD height="25" width="65%"> <font color="#000000"> 
            <input type="text" name="tel" maxlength="15" value=<%=tel%>>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><STRONG>请选择你所在地区: 
            </STRONG></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <SELECT name="province" size=1>
              <%if not province="" then%>
              <option value=<%=province%> selected><%=province%></option>
              <%end if%>
              <option value=""></option>
              <option value="北京">北京</option>
              <option value="广东">广东</option>
              <option value="广西">广西</option>
              <option value="海南">海南</option>
              <option value="福建">福建</option>
              <option value="天津">天津</option>
              <option value="湖南">湖南</option>
              <option value="湖北">湖北</option>
              <option value="河南">河南</option>
              <option value="河北">河北</option>
              <option value="山东">山东</option>
              <option value="山西">山西</option>
              <option value="黑龙江">黑龙江</option>
              <option value="辽宁">辽宁</option>
              <option value="上海">上海</option>
              <option value="甘肃">甘肃</option>
              <option value="青海">青海</option>
              <option value="新疆">新疆</option>
              <option value="西藏">西藏</option>
              <option value="宁夏">宁夏</option>
              <option value="四川">四川</option>
              <option value="云南">云南</option>
              <option value="吉林">吉林</option>
              <option value="内蒙古">内蒙古</option>
              <option value="陕西">陕西</option>
              <option value="安徽">安徽</option>
              <option value="贵州">贵州</option>
              <option value="江苏">江苏</option>
              <option value="重庆">重庆</option>
              <option value="浙江">浙江</option>
              <option value="江西">江西</option>
            </SELECT>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><STRONG>请选择你的性别:</STRONG></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <input <%if sex=0 then%>CHECKED<%end if%> name="sex" type=radio value=0>
            男 
            <input <%if sex=1 then%>checked<%end if%> name="sex"
      type=radio value=1>
            女 </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><B>出生日期：</B></font></TD>
          <TD width="65%"><font color="#000000"> 
            <input maxLength=4 name="year" size=4 value=<%=byear%>>
            年 
            <input maxLength=2 
      name="month" size=2 value=<%=bmonth%>>
            月 
            <input maxLength=2 name="day" size=2 value=<%=bday%>>
            日</font></TD>
        </TR>
        <tr> 
          <td align=left valign=center width="35%"><font color="#000000"><b>详细地址：</b></font></td>
          <td width="65%"> <font color="#000000"> 
            <input type="text" name="address" size="40" maxlength="100" value="<%=address%>">
            </font></td>
        </tr>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><B>教育背景：</B></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <select name="education" size=1>
              <option value="<%=education%>" selected><%=education%></option>
              <option value="大学">大学</option>
              <option value="专科">专科</option>
              <option value="高中">高中</option>
              <option value="硕士">硕士</option>
              <option value="博士">博士</option>
              <option value="其它">其它</option>
            </select>
            </font></TD>
        </TR>
        <TR> 
          <TD align=left vAlign=center width="35%"><font color="#000000"><STRONG>你的职业:</STRONG></font></TD>
          <TD width="65%"> <font color="#000000"> 
            <select name="vocation" size=1>
              <option value="<%=vocation%>" selected><%=vocation%></option>
              <option value="教育业">教育业</option>
              <option value="学生">学生</option>
              <option value="科研单位">科研单位</option>
              <option value="计算机业">计算机业</option>
              <option value="金融业">金融业</option>
              <option value="商业">商业</option>
              <option value="服务行业">服务行业</option>
              <option value="教育业">教育业</option>
              <option value="工程师">工程师</option>
              <option value="主管，经理">主管，经理</option>
              <option value="政府部门">政府部门</option>
              <option value="制造业">制造业</option>
              <option value="销售/广告/市场">销售/广告/市场</option>
              <option value="其它">其它</option>
            </select>
            </font></TD>
        </TR>
      </TBODY>
    </TABLE>
    <TABLE align=center bgColor=#DBE7F2 border=0 cellPadding=5 cellSpacing=2 
width="86%">
      <TBODY> 
      <TR> 
          <TD align=left vAlign=center width="35%"><STRONG><font color="#000000">感兴趣的学科领域： 
            （多选）</font></STRONG></TD>
        </TR>
  <TR> 
          <TD align=middle height="103"> 
		  <%showBack=things
		  if not isnull(showBack) then
		  showBack=replace(showBack,"生物学,","")
		  showBack=replace(showBack,"数学,","")
		  showBack=replace(showBack,"物理,","")
		  showBack=replace(showBack,"化学,","")
  		  showBack=replace(showBack,"经济,","")
 		  showBack=replace(showBack,"人工智能,","")
		  showBack=replace(showBack,"人工生命,","")
  		  showBack=replace(showBack,"社会学,","")
		  showBack=replace(showBack,"心理学,","")
		  showBack=replace(showBack,"管理学,","")
		  showBack=replace(showBack,"系统科学,","")
		  showBack=replace(showBack,"计算机科学,","")
		  end if
		  %>
            <table width="100%" border="0">
              <tr> 
                <td width="24%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"生物学")<>0 then%>
                  checked
                <%end if%> name="hobby" value="生物学">
                  生物学</font></td>
                <td width="23%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"数学")<>0 then%>
                  checked
                <%end if%> name="hobby" value="数学">
                  数学</font></td>
                <td width="26%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"物理")<>0 then%>
                  checked
                <%end if%> name="hobby" value="物理">
                  物理</font></td>
                <td width="27%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"化学")<>0 then%>
                  checked
                <%end if%> name="hobby" value="化学">
                  化学</font></td>
              </tr>
              <tr> 
                <td width="24%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"经济")<>0 then%>
                  checked
                <%end if%> name="hobby" value="经济">
                  经济</font></td>
                <td width="23%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"人工智能")<>0 then%>
                  checked
                <%end if%> name="hobby" value="人工智能">
                  人工智能</font></td>
                <td width="26%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"人工生命")<>0 then%>
                  checked
                <%end if%> name="hobby" value="人工生命">
                  人工生命</font></td>
                <td width="27%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"社会学")<>0 then%>
                  checked
                <%end if%> name="hobby" value="社会学">
                  社会学 </font></td>
              </tr>
              <tr> 
                <td width="24%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"心理学")<>0 then%>
                  checked
                <%end if%> name="hobby" value="心理学">
                  心理学</font></td>
                <td width="23%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"管理学")<>0 then%>
                  checked
                <%end if%> name="hobby" value="管理学">
                  管理学</font></td>
                <td width="26%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"系统科学")<>0 then%>
                  checked
                <%end if%> name="hobby" value="系统科学">
                  系统科学 </font></td>
                <td width="27%"> <font color="#000000"> 
                  <input type="checkbox"
                <%if instr(things,"计算机科学")<>0 then%>
                  checked
                <%end if%> name="hobby" value="计算机科学">
                  计算机</font></td>
              </tr>
            </table>
            <p>其它:<font color="#000000">
              <input name="otherback" type="text" id="otherback" value="<%=showBack%>" size="40" maxlength="100">
              （用,号隔开） </font></p>
            </TD>
      </TR></TBODY></TABLE><BR>
    <input type="submit" name="complete" value="确定">
    <input type="reset" name="Submit2" value="重写">
  </FORM>
</DIV>
</BODY></HTML>
