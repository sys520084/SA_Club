<%@ Language=VBScript %>
<!--#include file="../../var.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD><style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:����;color:black;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0000CC}
a:visited {text-decoration: none;; color: #660066}
a:hover {text-decoration: underline; color: #FF0000}
a:active {text-decoration;background-color: #CCCCCC; color: #006666}
-->
</style>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>��������Ա����</title>
<script language=javascript>
function convertdate(year1,month1,day1,year2,month2,day2,destdate1,destdate2){
t=true;
if(year1.value.length==0&&year2.value.length==0&&month1.value.length==0&&day1.value.length==0&&month2.value.length==0&&day2.value.length==0)
{
    destdate2.value="";
    destdate1.value="";
    return t;
}
t=isnumber(year1);
//alert(year1.value);
if(!t)return(t);
t=checkmonth(month1);
if(!t)return(t);
t=checkday(day1);
if(!t)return(t);
t=isnumber(year2);
if(!t)return(t);
t=checkmonth(month2);
if(!t)return(t);
t=checkday(day2);
if(!t)return(t);
  if(year1.value*10000+month1.value*100+day1.value>year2.value*10000+month2.value*100+day2.value){
    destdate2.value=year1.value+"-"+month1.value+"-"+day1.value;
    destdate1.value=year2.value+"-"+month2.value+"-"+day2.value;
  }else{
    destdate1.value=year1.value+"-"+month1.value+"-"+day1.value;
    destdate2.value=year2.value+"-"+month2.value+"-"+day2.value;
  }
//alert(destdate1.value);
//alert(destdate2.value);
return true;
}
function isnumber(v){
  if(!isNaN(v.value)&&v.value.length>0){
     return true;
  }else{   
     alert("����������Ч!");
     v.focus();
     return false;
  }
  return false;
}    
function checkmonth(v){
 if(v.value>12||v.value<0){
    alert("�·�������Ч��");
    v.focus();
    return false;
 }
 return isnumber(v);   
}
function checkday(v){
  if(v.value>31||v.value<0){
    alert("����������Ч");
    v.focus();
    return false;
   }
   return isnumber(v); 
}
</script></HEAD>
<BODY bgcolor="#FFFFFF">
<p align="center"> 
  <%dim catagory()
  dim ids()
if request("type")="ȷ��" then
    set conn=server.CreateObject("ADODB.Connection")
    conn.Open OpenString
    sql="select * from power where username='"&request("name")&"'"
    set rs=conn.Execute(sql)
    if rs.eof then
       Response.Write "<div algin='center'>�Բ��������ܴ򿪱�ҳ��<a href='javascript:history.go(-1)'>�����µ�½</a>"
       response.end
    end if
    if trim(rs("passwd"))<>request("pass") then   
       Response.Write "<div algin='center'>�Բ��������ܴ򿪱�ҳ��<a href='javascript:history.go(-1)'>�����µ�½</a>"
       Response.end
    end if
    if trim(rs("power"))<>2 then
       Response.Write "<div algin='center'>�Բ��������ܴ򿪱�ҳ��<a href='javascript:history.go(-1)'>�����µ�½</a>"
       response.end 
    end if
    session("admin")=true
    rs.close
    set rs=nothing
    sql="select id,catagory from forum_state"
    set rs=conn.Execute(sql)
    i=0
    do while not rs.eof
       i=i+1
       redim preserve catagory(i)
       redim preserve ids(i)
       catagory(i)=rs("catagory")
       id=cint(rs("id"))
       ids(i)=id
       rs.movenext
    loop   
    countforum=i
 %>
  <a href="banzhuapply.asp">�����������б�</a> </p>
<form action="sendsysmsg.asp" method="post" name="form5" target="_blank">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFCC">
    <tr> 
      <td><div align="center"> 
          <p><a href="sendsysmsg.asp" target="_blank">�鿴���͵�ϵͳ��Ϣ</a>&nbsp;&nbsp; ����ҷ�����Ϣ��</p>
          <p>��Ŀ��
            <input name="title" type="text" id="title">
          </p>
          <p>���ݣ� 
            <textarea name="content" id="content"></textarea>
          </p>
          <p>ָ�������ߣ�
            <input name="reciever" type="text" id="reciever">
          </p>
          <p> 
            <input type="submit" name="Submit10" value="ȷ��">
            <input type="reset" name="Submit11" value="��д">
          </p>
        </div></td>
    </tr>
  </table>
</form>
<p>&nbsp; </p>
<form name="form2" action="banzhu.asp" target="_blank" method="POST">
  <div align="center"> 
    <table width="100%" border="0" bgcolor="#CCCCFF">
      <tr align="center"> 
        <td>�� 
          <input type="text" name="name">
          Ϊ 
          <select name="catagory">
          <%for i=1 to countforum%>
            <option value="<%=catagory(i)%>"><%=catagory(i)%></option>
          <%next%>             
          </select>
          ����
<p> ���Ĺ���������(���Բ�ָ������ʹ�����������е����롣)��<br>
            <input type="password" name="pass">
            <br>
            <input type="submit" name="admin" value="ȷ��">
            <input type="reset" name="Submit3" value="��д">
          </p>
          <p> <a href="banzhu.asp" target="_blank" >�����б�</a></p>      </td>
      </tr>
    </table>
    
  </div>
</form>
<form method="post" action="addnewforum.asp" name="frmforum" target=blank>
  <table width="100%" border="0">
    <tr bgcolor="#FFCCCC"> 
      <td> 
        <div align="center">�����°棺 
          <input type="text" name="forum" size="20">
          ,����
          <select name="father">
            <option value="0"></option>
           <%for i=1 to countforum%>
            <option value="<%=ids(i)%>"><%=catagory(i)%></option>
          <%next%> 
          </select>
          ���Ӱ棬�Ƿ��Ǳ��⣿
          <input type="radio" name="title" value="0">
          �� 
          <input type="radio" name="title" value="1" checked>
          �� 
          <input type="submit" name="Submit7" value="ȷ��">
          <input type="reset" name="Submit9" value="��д">
        </div>
      </td>
    </tr>
  </table>
  </form>
<form name="form3" action="changepower.asp" target="_blank" method="POST" >
  <table width="100%" border="0">
    <tr bgcolor="#99CCFF"> 
      <td> 
        <p align="center">��� 
          <input type="text" name="username">
          �� 
          <select name="catagory">
            <%for i=1 to countforum%>
            <option value="<%=catagory(i)%>"><%=catagory(i)%></option>
          <%next%>  
          </select>
          ��Ȩ��</p>
        <p align="center"> 
          <input type="submit" name="Submit4" value="ȷ��">
          <input type="reset" name="Submit4" value="��д">
        </p>
        <p align="center"><a href="changepower.asp" target="_blank" >����Ȩ�����б�</a> 
        </p>
</td>
    </tr>
  </table>
  </form>
<form method="post" target="_blank" action="admin.asp">
  <div align="center">
    <table width="100%" border="0" bgcolor="#CCCCCC">
      <tr align="center"> 
        <td>�鿴 
          <select name="catagory">
            <%for i=1 to countforum%>
            <option value="<%=catagory(i)%>"><%=catagory(i)%></option>
          <%next%>  
          </select>
          ������� 
          <input type="submit" name="Submit" value="ȷ��">
        </td>
      </tr>
    </table>
    
  </div>
</form>
<div align="center">
  <p><b>�û�����</b> </p>
</div>
<form name="frm1" action="queryresult.asp" method="POST" target="_blank" onsubmit="return convertdate(this.timebyear,this.timebmonth,this.timebday,this.timebyear2,this.timebmonth2,this.timebday2,this.timehidebirthday1,timehidebirthday2)">
  <p align="center"><b>�û���ϸ��Ϣ��</b></p>
  <table align=center bgcolor=#FFCCCC border=0 cellpadding=5 cellspacing=2 
width="80%">
    <tr> 
      <td align=left valign=center width="35%"><b>�û�����</b></td>
      <td width="65%"> 
        <input type="text" name="username">
      </td>
    </tr>
   <tbody> 
    <tr> 
      <td align=left valign=center width="35%"><strong>����:</strong></td>
      <td width="65%"> 
        <input name="name" size=20 maxlength="20">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><strong>E-mail:</strong></td>
      <td width="65%"> 
        <input name="email" size=30 maxlength="30">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center 
width="35%"><strong>Homepage:</strong></td>
      <td width="65%"> 
        <input name="homepage" size=30 maxlength="50">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><b>�������������֤�ţ�</b></td>
      <td width="65%"> 
        <input type="text" name="idcard" maxlength="20">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><strong>���������OICQ:</strong></td>
      <td width="65%"> 
        <input name="oicq" maxlength="15">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%" height="25"> 
        <p><b>�绰��</b></p>
      </td>
      <td height="25" width="65%"> 
        <input type="text" name="tel" maxlength="10">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><b>BP�����룺</b></td>
      <td width="65%"> 
        <input type="text" name="bp" maxlength="20">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><strong>����ʡ��: </strong></td>
      <td width="65%"> 
        <select name="province" size=1>
          <option value="" selected></option>
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
        </select>
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><strong>�Ա�:</strong></td>
      <td width="65%"> 
        <input  name="radiosex" type=radio value=0>
        �� 
        <input  name="radiosex" type=radio value=1>
        Ů </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><b>�������</b></td>
      <td width="65%">
        <input maxlength=4 name="timebyear" size=4 >
        �� 
        <input maxlength=2 
      name="timebmonth" size=2 >
        �� 
        <input maxlength=2 name="timebday" size=2>
        �� - 
        <input maxlength=4 name="timebyear2" size=4 >
        �� 
        <input maxlength=2 
      name="timebmonth2" size=2 >
        �� 
        <input maxlength=2 name="timebday2" size=2>
        �� </td>
      <input type=hidden name=timehidebirthday1>
      <input type=hidden name=timehidebirthday2>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><b>��ϸ��ַ��</b></td>
      <td width="65%"> 
        <input type="text" name="likeaddress" size="40" maxlength="100">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><b>�������룺</b></td>
      <td width="65%"> 
        <input type="text" name="postno" maxlength="10">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><b>���ѧ��</b></td>
      <td width="65%"> 
        <select name="education" size=1>
          <option selected value=""></option>
          <option value=��ѧ>��ѧ</option>
          <option value=Сѧ>Сѧ</option>
          <option value=����>����</option>
          <option value=����>����</option>
          <option value=˶ʿ>˶ʿ</option>
          <option value=��ʿ>��ʿ</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><strong>ְҵ:</strong></td>
      <td width="65%"> 
        <select name="vocation" size=1>
          <option selected value=""></option>
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
        </select>
      </td>
    </tr>
    </tbody> 
  </table>  
  <table align=center bgcolor=#FFCCCC border=0 cellpadding=5 cellspacing=2 
width="80%">
    <tbody> 
    <tr> 
      <td align=left valign=center width="35%"><strong>��Ȥ���ã� ����ѡ��</strong></td>
    </tr>
    <tr> 
      <td align=middle height="103"> 
        <table width="100%" border="0">
          <tr> 
            <td width="22%"><b>���֣�</b></td>
            <td width="24%">&nbsp;</td>
            <td width="26%">&nbsp;</td>
            <td width="28%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="22%"> 
              <input type="checkbox"  name="checkinteremusic" value="pop">
              �������� </td>
            <td width="24%"> 
              <input type="checkbox" name="checkinteremusic" value="rock">
              ҡ������ </td>
            <td width="26%"> 
              <input type="checkbox" name="checkinteremusic" value="folk">
              �������� </td>
            <td width="28%"> 
              <input type="checkbox" name="checkinteremusic" value="english">
              ŷ���������� </td>
          </tr>
          <tr> 
            <td width="22%"> 
              <input type="checkbox" name="checkinteremusic" value="rihan">
              �պ����� </td>
            <td width="24%"> 
              <input type="checkbox" name="checkinteremusic" value="dance">
              �������� </td>
            <td width="26%"> 
              <input type="checkbox" name="checkinteremusic" value="classic">
              �����ŵ�����</td>
            <td width="28%">&nbsp;</td>
          </tr>
        </table>
        <table width="100%" border="0">
          <tr> 
            <td width="28%"><b>Ӱ�ӣ�</b></td>
            <td width="23%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
            <td width="24%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="28%"> 
              <input type="checkbox" name="checkinteremovie" value="entertainment">
              ��ҵ����Ƭ </td>
            <td width="23%"> 
              <input type="checkbox" name="checkinteremovie" value="art">
              ��������Ƭ </td>
            <td width="25%"> 
              <input type="checkbox" name="checkinteremovie" value="discovery">
              ��¼̽��Ƭ </td>
            <td width="24%"> 
              <input type="checkbox" name="checkinteremovie" value="cartoon">
              ������ͨƬ </td>
          </tr>
        </table>
      </td>
    </tr>
    </tbody> 
  </table>
  <p align="center"> 
    <input type="hidden" name="hidetable" value="detail">
    <input type="submit" name="buttonok" value="ȷ��">
    <input type="reset" name="buttonSubmit6" value="��д">
  </p>
</form>
<form method="post" action="queryresult.asp" name="frm2" target="_blank" onsubmit="return convertdate(this.timebyear,this.timebmonth,this.timebday,this.timebyear2,this.timebmonth2,this.timebday2,this.timehidesigndate1,timehidesigndate2)">
    <p align="center"><b>�����û���Ϣ��</b></p>
    <table width="80%" border="0" align="center" bgcolor="#FFCCFF">
      <tr> 
        <td width="35%"><b>�û�����</b></td>
        <td width="65%"> 
          <input type="text" name="username">
        </td>
      </tr>
      <tr> 
        <td width="35%"><b>�Ǽ����ڣ�</b></td>
        <td width="65%"> 
          <input maxlength=4 name="timebyear" size=4 value=<%=year(date)%>>
          �� 
          <input maxlength=2 name="timebmonth" size=2 value=<%=month(date)%>>
          �� 
          <input maxlength=2 name="timebday" size=2 value=<%=day(date)%>>
          �� - 
          <input maxlength=4 name="timebyear2" size=4 value=<%=year(date)%>>
          �� 
          <input maxlength=2 name="timebmonth2" size=2 value=<%=month(date)%>>
          �� 
          <input maxlength=2 name="timebday2" size=2 value=<%=day(date)%>>
          �� 
          
        <input type=hidden name=timehidesigndate1>
          
        <input type=hidden name=timehidesigndate2>
      </td>
      </tr>
    </table>
    <p align="center">
      <input type="hidden" name="hidetable" value="users">
      <input type="submit" name="buttonok" value="ȷ��">
      <input type="reset" name="Submit8" value="��д">
    </p>
  </form>
<p>&nbsp;</p><form name="form4" action="sqlself.asp" target="_blank">
  <div align="center">
    <table width="100%" border="0" bgcolor="#CCCCCC">
      <tr align="center"> 
        <td>�Լ�����SQL��䣺<br>
          <textarea name="sql" cols="80"></textarea>
          <br>
          <input type="submit" name="Submit5" value="ȷ��">
          <input type="reset" name="Submit6" value="��д">
        </td>
      </tr>
    </table>
    
  </div>
</form>
<p> <%else%> </p>
<p>&nbsp;</p><form method="post" action="jingwen.asp" id=form1 name=form1>
  <p align="center">����Ա���� 
    <input type="text" name="name">
  </p>
  <p align="center">���룺&nbsp;&nbsp; 
    <input type="password" name="pass">
  </p>
  <p align="center"> 
    <input type="submit" name="type" value="ȷ��">
    <input type="reset" name="Submit2" value="��д">
  </p>
</form>
<%end if%>
</BODY>
</HTML>
