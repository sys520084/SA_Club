<!--#include file="../../var.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>���²���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:����;color:black;}
a:link {text-decoration: none; ; font-size: 9pt; color: #0000CC}
a:visited {text-decoration: none;; color: #660066}
a:hover {text-decoration: underline; color: #FF0000}
a:active {text-decoration;background-color: #CCCCCC; color: #006666}
.unnamed1 {  font-family: "����"; font-size: 9pt; color: #000000}
-->
</style>
<script language="javascript">
function checkout(){
  if(form1.yue.value!=""&&form1.yue1.value!=""&&form1.ri.value!=""&&form1.ri1.value!=""){
     form1.fromdate.value=form1.nian.value+"-"+form1.yue.value+"-"+form1.ri.value;
     form1.todate.value=form1.nian1.value+"-"+form1.yue1.value+"-"+form1.ri1.value;
   }  
 form1.submit();  
}
</script>
</head>

<body bgcolor="#FFFFFF">
<div align="center">
  <p class="unnamed1">��������ҵ�������</p>
  <form name="form1" action="viewtitle.asp" method="post" onsubmit="checkout();">
    <p class="unnamed1"> �����а������֣� 
      <input type="text" name="title" value="">
    </p>
    <p class="unnamed1"> ���������ڣ� 
      <input type=hidden value="" name="fromdate">
      <input type=hidden value="" name="todate">
      <select name="nian">
	  <%yea=year(date)
	  for i=2000 to yea%>
        <option value=<%=i%>><%=i%></option>
      <%next%>
      </select>
      �� 
      <select name="yue">
        <option value=""></option>
        <option value="01">1</option>
        <option value="02">2</option>
        <option value="03">3</option>
        <option value="04">4</option>
        <option value="05">5</option>
        <option value="06">6</option>
        <option value="07">7</option>
        <option value="08">8</option>
        <option value="09">9</option>
        <option value="10">10</option>
        <option value="11">11</option>
        <option value="12">12</option>
      </select>
      �� 
      <select name="ri">
        <option value=""></option>
        <option value="01">1</option>
        <option value="02">2</option>
        <option value="03">3</option>
        <option value="04">4</option>
        <option value="05">5</option>
        <option value="06">6</option>
        <option value="07">7</option>
        <option value="08">8</option>
        <option value="09">9</option>
        <option value="10">10</option>
        <option value="11">11</option>
        <option value="12">12</option>
        <option value="13">13</option>
        <option value="14">14</option>
        <option value="15">15</option>
        <option value="16">16</option>
        <option value="17">17</option>
        <option value="18">18</option>
        <option value="19">19</option>
        <option value="20">20</option>
        <option value="21">21</option>
        <option value="22">22</option>
        <option value="23">23</option>
        <option value="24">24</option>
        <option value="25">25</option>
        <option value="26">26</option>
        <option value="27">27</option>
        <option value="28">28</option>
        <option value="29">29</option>
        <option value="30">30</option>
        <option value="31">31</option>
      </select>
      �գ�</p>
    <p class="unnamed1">�� 
      <select name="nian1">
        <%for i=2000 to yea%>
        <option value=<%=i%>><%=i%></option>
		<%next%>
      </select>
      �� 
      <select name="yue1">
        <option value=""></option>
        <option value="01">1</option>
        <option value="02">2</option>
        <option value="03">3</option>
        <option value="04">4</option>
        <option value="05">5</option>
        <option value="06">6</option>
        <option value="07">7</option>
        <option value="08">8</option>
        <option value="09">9</option>
        <option value="10">10</option>
        <option value="11">11</option>
        <option value="12">12</option>
      </select>
      �� 
      <select name="ri1">
        <option value=""></option>
        <option value="01">1</option>
        <option value="02">2</option>
        <option value="03">3</option>
        <option value="04">4</option>
        <option value="05">5</option>
        <option value="06">6</option>
        <option value="07">7</option>
        <option value="08">8</option>
        <option value="09">9</option>
        <option value="10">10</option>
        <option value="11">11</option>
        <option value="12">12</option>
        <option value="13">13</option>
        <option value="14">14</option>
        <option value="15">15</option>
        <option value="16">16</option>
        <option value="17">17</option>
        <option value="18">18</option>
        <option value="19">19</option>
        <option value="20">20</option>
        <option value="21">21</option>
        <option value="22">22</option>
        <option value="23">23</option>
        <option value="24">24</option>
        <option value="25">25</option>
        <option value="26">26</option>
        <option value="27">27</option>
        <option value="28">28</option>
        <option value="29">29</option>
        <option value="30">30</option>
        <option value="31">31</option>
      </select>
      ��֮�� </p>
    <p class="unnamed1">�����ǣ� 
	  <input name="authortype" type="hidden" value=1>
      <input name="author" type="text" id="author" value="">
    </p>
    <p class="unnamed1">���°������֣� 
      <input name="content" type="text" id="content" value="">
    </p>
    <p class="unnamed1">���棺 
      <%set conn=server.CreateObject("ADODB.Connection")
	Conn.open OpenString
    set rs=server.CreateObject("ADODB.Recordset")
	dim sql
	sql="select * from forum_state"
	rs.open sql,conn,3
'	response.end
	%>
      <select name="Catagory" id="Catagory">
        <option value="">ȫ��</option>
        <%do while not rs.eof%>
        <option value="<%=rs("catagory")%>" <%if rs("catagory")=request("catagory") then%>selected<%end if%>><%=rs("catagory")%></option>
        <%rs.movenext
	  loop%>
      </select>
    </p>
    <p class="unnamed1">�Ƿ���һ����� 
      <input name="subs" type="radio" value="yes">
      ��&nbsp;&nbsp;&nbsp;&nbsp; 
      <input  name="subs" type="radio" value="no" checked>
      �� </p>
    <p> 
      <input type="submit" name="finder" value="ȷ��">
      <input type="reset" name="Submit2" value="��д">
    </p>
  </form>
  <p>&nbsp;</p>
</div>
</body>
</html>
