<%@ Language=VBScript %>
<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<%OpenString=DBBS()
detailcontent=""
function checkvalid(txt,tp)
     if len(txt)>0 then
        for i=1 to len(txt)
		   ass=asc(mid(txt,i,1))
		   if tp=0 then
			   if (ass<48 and ass>0) or (ass>57 and ass<65) or ass>122  then
				  Response.Write "<div align=center>对不起，用户名中只能含有字符或者数字！<a href='javascript:history.go(-1);'>请重试</a></div>"
				  Response.End 
			   end if
		   else
				if ass="'" then
					Response.Write "<div align=center>对不起，口令中不能包含<font color='red'>'</font>字符！<a href='javascript:history.go(-1);'>请重试</a></div>"
					Response.End 
				end if
		   end if
        next 
     end if
	 checkvalid=txt
end function 
%>
<HTML>
<HEAD>
<STYLE type=text/css>
DIV,p {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; color:#000000;}
TD {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
}
A {COLOR: #0099FF; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #FF0000; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
input,textarea,SELECT {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	color: #000000;
	background-color: #FFFFFF;
}</STYLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</HEAD>
<BODY bgcolor="#FFFFFF">
<P>
<%
'response.write request("username")
     if isnull(request("username")) or request("username")="" then
        Response.Write "<div align='center'>错误！用户名是空的！</div>"
        Response.End
     end if
     set conn=server.CreateObject("ADODB.Connection")
     conn.Open OpenString
     conn.BeginTrans 
     set rs=server.CreateObject("ADODB.Recordset")
     'Response.Write request("type")
     if request("type")="new" then
		username=checkvalid(request("username"),0)
		rs.open "select * from users where username='"&username&"'",conn,3
		if not rs.eof then%>
			<DIV align=center>对不起，<font color="red"><%=username%></font>这个名字已经有人申请了！</DIV>
		<%response.end			
		end if
		rs.close
        rs.Open "users",conn,1,3
        rs.AddNew
        conn.Errors.clear
        'on error resume next
        rs("username")=username
        rs("passwd")=checkvalid(request("pass"),1)
        rs("nickname")=MyhtmlEncode(request("nickname"))
        rs("signdate")=date
        rs("lastdate")=date
        rs("times")=0
        rs("markscore")=0
        rs("articles")=0
        'on error resume next
        rs.Update
        if conn.Errors.Count>0 then
           for i=0 to conn.Errors.Count-1
               Response.Write conn.Errors(i).Description
           next
		   response.write "error"
           Response.end
        end if       
        rs.Close
        rs.Open "detail",conn,1,3
        rs.AddNew
        rs("username")=username
        rs("name")=MyhtmlEncode(request("truename"))
        rs("question")=MyhtmlEncode(request("question"))
        rs("anwser")=MyhtmlEncode(request("anwser"))
        rs("sex")=cint(request("sex"))
        rs("homepage")=MyhtmlEncode(request("homepage"))
        rs("email")=MyhtmlEncode(request("email"))
        rs("oicq")=MyhtmlEncode(request("oicq"))
		rs("msn")=MyhtmlEncode(request("msn"))
        rs("tel")=MyhtmlEncode(request("tel"))
		'response.write date(request("year")&"-"&request("month")&"-"&request("day"))
		'response.end
		if request("year")="" then 
			year1=0
		else
			year1=cint(request("year"))
		end if
		if request("month")="" then
			month1=0
		else
			month1=cint(request("month"))
		end if
		if request("day")="" then
			day1=0
		else
			day1=cint(request("day"))
		end if
		if year1>1900 and month1>0 and month1<12 and day1>0 and day1<31 then
			rs("birthday")=DateSerial(year1,month1,day1)
		end if
        rs("province")=MyhtmlEncode(request("province"))
        rs("address")=MyhtmlEncode(request("address"))
        rs("education")=MyhtmlEncode(request("education"))
        rs("vocation")=MyhtmlEncode(request("vocation"))
        if isnull(request("hobby")) or request("hobby")="" then
           ihobby=""
        else   
           for j=1 to Request.Form("hobby").Count 
                ihobby=ihobby&Request.Form("hobby")(j)&","
           next
        end if       
        rs("interething")=ihobby&replace(replace(request("otherback"),"，",",")," ",",")
        rs.Update
        session("user")=username
		randomize
		face=int(rnd()*10)+1
		'response.write face
        sql="insert into memberdes (username,showtitle,detailcontent,portrait) values ('"&username&"','"&username&"的头脑风暴','"&detailcontent&"','"&"pre\"&face&".jpg"&"')"
        conn.Execute(sql)
     else
        conn.Errors.clear
		username=request("username")
        'on error resume next
        if session("user")="" then
			Response.Write "<div align='center'>对不起您不能更改用户信息，请<a href='relogin.asp'>重新登陆</a>。</div>"
			Response.End
		end if
        'on error resume next

        rs.Open "select nickname from users where username='"&username&"'",conn,1,3
        rs("nickname")=MyhtmlEncode(request("nickname"))
        rs.Update
        rs.Close
        rs.Open "select * from detail where username='"&username&"'",conn,1,3
        rs("name")=MyhtmlEncode(request("truename"))
        rs("question")=MyhtmlEncode(request("question"))
        rs("anwser")=MyhtmlEncode(request("anwser"))
        rs("sex")=cint(request("sex"))
        rs("homepage")=MyhtmlEncode(request("homepage"))
        rs("email")=MyhtmlEncode(request("email"))
        rs("oicq")=MyhtmlEncode(request("oicq"))
		rs("msn")=MyhtmlEncode(request("msn"))
        rs("tel")=MyhtmlEncode(request("tel"))
		myyear=request("year")
		mymonth=request("month")
		myday=request("day")
		if isnull(myyear) or myyear="" then
			myyear="1900"
		end if
		if isnull(mymonth) or mymonth="" then
			mymonth="1"
		end if
		if isnull(myday) or myday="" then
			myday="1"
		end if
		year1=cint(myyear)
		month1=cint(mymonth)
		day1=cint(myday)
		'response.end 
		if year1>=1900 and month1>0 and month1<=12 and day1>0 and day1<=31 then
			rs("birthday")=DateSerial(year1,month1,day1)
		end if
        rs("province")=MyhtmlEncode(request("province"))
        rs("address")=request("address")
        rs("education")=MyhtmlEncode(request("education"))
        rs("vocation")=MyhtmlEncode(request("vocation"))
        if isnull(request("hobby")) or request("hobby")="" then
           ihobby=""
        else   
           for j=1 to Request.Form("hobby").Count 
                ihobby=ihobby&Request.Form("hobby")(j)&","
           next
        end if       
        rs("interething")=ihobby&replace(replace(request("otherback"),"，",",")," ",",")
		
        rs.Update
      end if      
      rs.Close
      if conn.Errors.Count>0 then
         Response.Write "<div align=center>对不起，您填入的数据有错误，可能是您的用户名为空或者出生日期有错，注册失败！</div>"
         Response.End
         conn.RollbackTrans
      end if    
      conn.CommitTrans  
 if request("type")="new" then %> </P>
<div align="center"> 
  <p>您已经成功的加入了头脑风暴社区,请继续进行个人说明设定或<a href="../forum.asp?user=<%=session("user")%>">直接进入社区</a>：</p>
  <p>也可以更改<strong><a href="persenal.asp">个人说明</a></strong></p>
  
<p>&nbsp;</p></div>
 
<div align="center"> </div>
 <%else%>
 <div align="center">您的资料修改成功，请继续进行操作！</div>
 <%end if%>  
</BODY>
</HTML>
