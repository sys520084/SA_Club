<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/define.css" rel="stylesheet" type="text/css" />
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:"宋体";color:#333333;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0066FF}
a:visited {text-decoration: none;; color: #0033CC}
a:hover {text-decoration: underline; color: #0033CC}
a:active {text-decoration;color: #0066FF}
.art {
	font-size: 10pt;
	color: #0000CC;
	font-style: normal;
}
-->
</style>

<script language="javascript">
<!--
function js_callpages(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=330");
  newwin.focus();
  return false;
}
-->
</script>
</head>
<body bgcolor="#FFFFFF">
<div align=left>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="left" valign="top" ><!--#include file="userdisplay.asp"-->
		  <%
  username=request("user")
   sql="select users.nickname as nickname,users.articles as articlecount,users.times as cometimes,users.signdate as signdate,users.markscore as markscore,detail.name as name,detail.sex as sex,detail.email as email,detail.homepage as homepage,detail.birthday as birthday, detail.interething as interests,"
  sql=sql&"memberdes.readme as readme,memberdes.showtitle as showtitle,memberdes.portrait as portrait,memberdes.detailcontent as detailcontent from users,detail,memberdes where users.username=detail.username and "
  sql=sql&"users.username=memberdes.username and users.username='"&username&"'"
  set rs=server.CreateObject("ADODB.Recordset")
  'Response.Write sql
  'Response.End
  rs.Open sql,conn
  if rs.EOF then
     Response.Write "<div align=center>对不起，没有找到您指定的用户，可能他(她)还没有在此出登记。</div>"
     Response.End
  end if   
  'on error resume next
  birthday=year(rs("birthday"))&"-"&month(rs("birthday"))
  sex=""
       if cint(rs("sex"))=0 then
          sex="男"
       end if
       if cint(rs("sex"))=1 then
          sex="女"
       end if
       story=MyHtmlEncode(trim(rs("readme")))
	   if rs("nickname")<>"" then
	    nickname="("&rs("nickname")&")"
	   else
	   	nickname=""
	  end if 
	portrait=rs("portrait")
       '计算经验值
      articlecount=cint(rs("articlecount"))
      cometimes=cint(rs("cometimes"))
      signdate=rs("signdate")
      markscore=cint(rs("markscore"))
      days=DateDiff("d",signdate,date)
      days=cint(days)
      if articlecount="" then
         articlecount=0
      end if
      if cometimes="" then
         cometimes=0
      end if
      if markscore="" then
         markscore=0
      end if         
      if days=0 then
         frequency=0
      else
         frequency=cometimes/days
      end if      
      value=articlecount+markscore*5+cometimes/4+frequency/2
      rank=value/500
	  nickname=rs("nickname")
	  if len(nickname)=0 or isnull(nickname) then
	  	nickname="无"
	  end if
	  interests=rs("interests")
	  if interests="" or isnull(interests) then
	  	interests="无"
	  end if
	  

  if not rs.eof then
  	'response.write rs("detailcontent")
  	detailcontent=rs("detailcontent")
	content=detailcontent
	'Dim regEx,strOut
  	Set regEx1 = New RegExp
  	regEx1.Pattern = "<.*?>"
  	regEx1.IgnoreCase = True 
  	regEx1.Global = True 
  	strOut = regEx1.Replace(detailcontent, "")
  	strOut=replace(strOut,"&nbsp;"," ")
  	strOut=trim(strOut)
	detailcontent=strOut
	'response.write content
	detailcontent=replace(detailcontent,chr(9),"")
	'detailcontent=replace(detailcontent,chr(13),"")
	'detailcontent=replace(detailcontent,chr(11),"")
	detailcontent=replace(detailcontent,vbTab,"")
	detailcontent=replace(detailcontent,vbCrlf,"")

	'detailcontent=replace(detailcontent,chr(11),"")

	'response.write content
	'response.end 
	if left(detailcontent,len("http://"))="http://" then
		response.write detailcontent
		'response.end 
		'response.redirect trim(detailcontent)
		%>
		<script language="javascript">
		window.open('<%=detailcontent%>');
		history.go(-1);
		</script>
		<%
		response.end
	end if
  %>
        <table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" bordercolor="#3366FF">
          <tr>
            <td><table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" bordercolor="#8894D7">
                <tr> 
                  <td align="left">&nbsp; </td>
                </tr>
                <tr> 
                  <td align="left"><%=content%></td>
                </tr>
                <tr> 
                  <td align="left">&nbsp;</td>
                </tr>
              </table></td>
            <td width="200" align="left" valign="top">
			<br>
              <p>
                <strong>昵称</strong>：<%=nickname%></p>
              <p> <strong>电子邮件：</strong><a href="mailto:<%=rs("email")%>"><%=rs("email")%></a></p>
              <p> <strong>感兴趣领域： </strong><%=interests%></p>
              <p> 
                <%if rs("homepage")<>"" then%>
                <strong>个人主页</strong>：<a href="<%=rs("homepage")%>"><%=rs("homepage")%></a> </p>
              <p> 
                <%end if%>
              </p>
			  <%if session("user")=username then%>
                <a href="memeber.asp?type=%22change%22"><img src="../images/edit.gif" width="20" height="20" alt="编辑修改用户信息" border="0"> 
                </a><br>
                <%end if%></td>
          </tr>
        </table> 
        <%end if
rs.close
set rs=nothing
%>
      </td>
	  </tr>  
	  </table>
  <p>&nbsp;</p>
</div>
</body>
</html>
