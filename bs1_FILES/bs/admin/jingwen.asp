<%@ Language=VBScript %>
<!--#include file="../../var.asp"-->
<%OpenString=DBBS()%>
<HTML>
<HEAD><style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:宋体;color:black;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0000CC}
a:visited {text-decoration: none;; color: #660066}
a:hover {text-decoration: underline; color: #FF0000}
a:active {text-decoration;background-color: #CCCCCC; color: #006666}
-->
</style>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>超级管理员管理</title>
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
     alert("日期输入无效!");
     v.focus();
     return false;
  }
  return false;
}    
function checkmonth(v){
 if(v.value>12||v.value<0){
    alert("月份输入无效！");
    v.focus();
    return false;
 }
 return isnumber(v);   
}
function checkday(v){
  if(v.value>31||v.value<0){
    alert("日期输入无效");
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
if request("type")="确定" then
    set conn=server.CreateObject("ADODB.Connection")
    conn.Open OpenString
    sql="select * from power where username='"&request("name")&"'"
    set rs=conn.Execute(sql)
    if rs.eof then
       Response.Write "<div algin='center'>对不起，您不能打开本页。<a href='javascript:history.go(-1)'>请重新登陆</a>"
       response.end
    end if
    if trim(rs("passwd"))<>request("pass") then   
       Response.Write "<div algin='center'>对不起，您不能打开本页。<a href='javascript:history.go(-1)'>请重新登陆</a>"
       Response.end
    end if
    if trim(rs("power"))<>2 then
       Response.Write "<div algin='center'>对不起，您不能打开本页。<a href='javascript:history.go(-1)'>请重新登陆</a>"
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
  <a href="banzhuapply.asp">版主申请人列表</a> </p>
<form action="sendsysmsg.asp" method="post" name="form5" target="_blank">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFCC">
    <tr> 
      <td><div align="center"> 
          <p><a href="sendsysmsg.asp" target="_blank">查看回送的系统消息</a>&nbsp;&nbsp; 给大家发送消息：</p>
          <p>题目：
            <input name="title" type="text" id="title">
          </p>
          <p>内容： 
            <textarea name="content" id="content"></textarea>
          </p>
          <p>指定接收者：
            <input name="reciever" type="text" id="reciever">
          </p>
          <p> 
            <input type="submit" name="Submit10" value="确定">
            <input type="reset" name="Submit11" value="重写">
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
        <td>封 
          <input type="text" name="name">
          为 
          <select name="catagory">
          <%for i=1 to countforum%>
            <option value="<%=catagory(i)%>"><%=catagory(i)%></option>
          <%next%>             
          </select>
          斑竹，
<p> 他的管理密码是(可以不指定，则使用他在社区中的密码。)：<br>
            <input type="password" name="pass">
            <br>
            <input type="submit" name="admin" value="确定">
            <input type="reset" name="Submit3" value="重写">
          </p>
          <p> <a href="banzhu.asp" target="_blank" >斑竹列表</a></p>      </td>
      </tr>
    </table>
    
  </div>
</form>
<form method="post" action="addnewforum.asp" name="frmforum" target=blank>
  <table width="100%" border="0">
    <tr bgcolor="#FFCCCC"> 
      <td> 
        <div align="center">增加新版： 
          <input type="text" name="forum" size="20">
          ,它是
          <select name="father">
            <option value="0"></option>
           <%for i=1 to countforum%>
            <option value="<%=ids(i)%>"><%=catagory(i)%></option>
          <%next%> 
          </select>
          的子版，是否是标题？
          <input type="radio" name="title" value="0">
          是 
          <input type="radio" name="title" value="1" checked>
          否 
          <input type="submit" name="Submit7" value="确定">
          <input type="reset" name="Submit9" value="重写">
        </div>
      </td>
    </tr>
  </table>
  </form>
<form name="form3" action="changepower.asp" target="_blank" method="POST" >
  <table width="100%" border="0">
    <tr bgcolor="#99CCFF"> 
      <td> 
        <p align="center">封闭 
          <input type="text" name="username">
          在 
          <select name="catagory">
            <%for i=1 to countforum%>
            <option value="<%=catagory(i)%>"><%=catagory(i)%></option>
          <%next%>  
          </select>
          的权限</p>
        <p align="center"> 
          <input type="submit" name="Submit4" value="确定">
          <input type="reset" name="Submit4" value="重写">
        </p>
        <p align="center"><a href="changepower.asp" target="_blank" >被封权限人列表</a> 
        </p>
</td>
    </tr>
  </table>
  </form>
<form method="post" target="_blank" action="admin.asp">
  <div align="center">
    <table width="100%" border="0" bgcolor="#CCCCCC">
      <tr align="center"> 
        <td>查看 
          <select name="catagory">
            <%for i=1 to countforum%>
            <option value="<%=catagory(i)%>"><%=catagory(i)%></option>
          <%next%>  
          </select>
          版的文章 
          <input type="submit" name="Submit" value="确定">
        </td>
      </tr>
    </table>
    
  </div>
</form>
<div align="center">
  <p><b>用户查找</b> </p>
</div>
<form name="frm1" action="queryresult.asp" method="POST" target="_blank" onsubmit="return convertdate(this.timebyear,this.timebmonth,this.timebday,this.timebyear2,this.timebmonth2,this.timebday2,this.timehidebirthday1,timehidebirthday2)">
  <p align="center"><b>用户详细信息：</b></p>
  <table align=center bgcolor=#FFCCCC border=0 cellpadding=5 cellspacing=2 
width="80%">
    <tr> 
      <td align=left valign=center width="35%"><b>用户名：</b></td>
      <td width="65%"> 
        <input type="text" name="username">
      </td>
    </tr>
   <tbody> 
    <tr> 
      <td align=left valign=center width="35%"><strong>姓名:</strong></td>
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
      <td align=left valign=center width="35%"><b>请输入您的身份证号：</b></td>
      <td width="65%"> 
        <input type="text" name="idcard" maxlength="20">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><strong>请输入你的OICQ:</strong></td>
      <td width="65%"> 
        <input name="oicq" maxlength="15">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%" height="25"> 
        <p><b>电话：</b></p>
      </td>
      <td height="25" width="65%"> 
        <input type="text" name="tel" maxlength="10">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><b>BP机号码：</b></td>
      <td width="65%"> 
        <input type="text" name="bp" maxlength="20">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><strong>所在省份: </strong></td>
      <td width="65%"> 
        <select name="province" size=1>
          <option value="" selected></option>
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
        </select>
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><strong>性别:</strong></td>
      <td width="65%"> 
        <input  name="radiosex" type=radio value=0>
        男 
        <input  name="radiosex" type=radio value=1>
        女 </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><b>出生年份</b></td>
      <td width="65%">
        <input maxlength=4 name="timebyear" size=4 >
        年 
        <input maxlength=2 
      name="timebmonth" size=2 >
        月 
        <input maxlength=2 name="timebday" size=2>
        日 - 
        <input maxlength=4 name="timebyear2" size=4 >
        年 
        <input maxlength=2 
      name="timebmonth2" size=2 >
        月 
        <input maxlength=2 name="timebday2" size=2>
        日 </td>
      <input type=hidden name=timehidebirthday1>
      <input type=hidden name=timehidebirthday2>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><b>详细地址：</b></td>
      <td width="65%"> 
        <input type="text" name="likeaddress" size="40" maxlength="100">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><b>邮政编码：</b></td>
      <td width="65%"> 
        <input type="text" name="postno" maxlength="10">
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><b>最高学历</b></td>
      <td width="65%"> 
        <select name="education" size=1>
          <option selected value=""></option>
          <option value=大学>大学</option>
          <option value=小学>小学</option>
          <option value=初中>初中</option>
          <option value=高中>高中</option>
          <option value=硕士>硕士</option>
          <option value=博士>博士</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td align=left valign=center width="35%"><strong>职业:</strong></td>
      <td width="65%"> 
        <select name="vocation" size=1>
          <option selected value=""></option>
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
        </select>
      </td>
    </tr>
    </tbody> 
  </table>  
  <table align=center bgcolor=#FFCCCC border=0 cellpadding=5 cellspacing=2 
width="80%">
    <tbody> 
    <tr> 
      <td align=left valign=center width="35%"><strong>兴趣爱好： （多选）</strong></td>
    </tr>
    <tr> 
      <td align=middle height="103"> 
        <table width="100%" border="0">
          <tr> 
            <td width="22%"><b>音乐：</b></td>
            <td width="24%">&nbsp;</td>
            <td width="26%">&nbsp;</td>
            <td width="28%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="22%"> 
              <input type="checkbox"  name="checkinteremusic" value="pop">
              流行音乐 </td>
            <td width="24%"> 
              <input type="checkbox" name="checkinteremusic" value="rock">
              摇滚音乐 </td>
            <td width="26%"> 
              <input type="checkbox" name="checkinteremusic" value="folk">
              民族音乐 </td>
            <td width="28%"> 
              <input type="checkbox" name="checkinteremusic" value="english">
              欧美流行音乐 </td>
          </tr>
          <tr> 
            <td width="22%"> 
              <input type="checkbox" name="checkinteremusic" value="rihan">
              日韩音乐 </td>
            <td width="24%"> 
              <input type="checkbox" name="checkinteremusic" value="dance">
              舞曲音乐 </td>
            <td width="26%"> 
              <input type="checkbox" name="checkinteremusic" value="classic">
              西方古典音乐</td>
            <td width="28%">&nbsp;</td>
          </tr>
        </table>
        <table width="100%" border="0">
          <tr> 
            <td width="28%"><b>影视：</b></td>
            <td width="23%">&nbsp;</td>
            <td width="25%">&nbsp;</td>
            <td width="24%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="28%"> 
              <input type="checkbox" name="checkinteremovie" value="entertainment">
              商业娱乐片 </td>
            <td width="23%"> 
              <input type="checkbox" name="checkinteremovie" value="art">
              人文艺术片 </td>
            <td width="25%"> 
              <input type="checkbox" name="checkinteremovie" value="discovery">
              记录探索片 </td>
            <td width="24%"> 
              <input type="checkbox" name="checkinteremovie" value="cartoon">
              动画卡通片 </td>
          </tr>
        </table>
      </td>
    </tr>
    </tbody> 
  </table>
  <p align="center"> 
    <input type="hidden" name="hidetable" value="detail">
    <input type="submit" name="buttonok" value="确定">
    <input type="reset" name="buttonSubmit6" value="重写">
  </p>
</form>
<form method="post" action="queryresult.asp" name="frm2" target="_blank" onsubmit="return convertdate(this.timebyear,this.timebmonth,this.timebday,this.timebyear2,this.timebmonth2,this.timebday2,this.timehidesigndate1,timehidesigndate2)">
    <p align="center"><b>基本用户信息：</b></p>
    <table width="80%" border="0" align="center" bgcolor="#FFCCFF">
      <tr> 
        <td width="35%"><b>用户名：</b></td>
        <td width="65%"> 
          <input type="text" name="username">
        </td>
      </tr>
      <tr> 
        <td width="35%"><b>登记日期：</b></td>
        <td width="65%"> 
          <input maxlength=4 name="timebyear" size=4 value=<%=year(date)%>>
          年 
          <input maxlength=2 name="timebmonth" size=2 value=<%=month(date)%>>
          月 
          <input maxlength=2 name="timebday" size=2 value=<%=day(date)%>>
          日 - 
          <input maxlength=4 name="timebyear2" size=4 value=<%=year(date)%>>
          年 
          <input maxlength=2 name="timebmonth2" size=2 value=<%=month(date)%>>
          月 
          <input maxlength=2 name="timebday2" size=2 value=<%=day(date)%>>
          日 
          
        <input type=hidden name=timehidesigndate1>
          
        <input type=hidden name=timehidesigndate2>
      </td>
      </tr>
    </table>
    <p align="center">
      <input type="hidden" name="hidetable" value="users">
      <input type="submit" name="buttonok" value="确定">
      <input type="reset" name="Submit8" value="重写">
    </p>
  </form>
<p>&nbsp;</p><form name="form4" action="sqlself.asp" target="_blank">
  <div align="center">
    <table width="100%" border="0" bgcolor="#CCCCCC">
      <tr align="center"> 
        <td>自己输入SQL语句：<br>
          <textarea name="sql" cols="80"></textarea>
          <br>
          <input type="submit" name="Submit5" value="确定">
          <input type="reset" name="Submit6" value="重写">
        </td>
      </tr>
    </table>
    
  </div>
</form>
<p> <%else%> </p>
<p>&nbsp;</p><form method="post" action="jingwen.asp" id=form1 name=form1>
  <p align="center">管理员名： 
    <input type="text" name="name">
  </p>
  <p align="center">密码：&nbsp;&nbsp; 
    <input type="password" name="pass">
  </p>
  <p align="center"> 
    <input type="submit" name="type" value="确定">
    <input type="reset" name="Submit2" value="重写">
  </p>
</form>
<%end if%>
</BODY>
</HTML>
