<!--#include file="../../../VAR.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>����������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:black;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
-->
</style>
<script language="javascript">
<!--
function js_callpage(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=330");
  newwin.focus();
  return false;
}
function move(loc)
{
	var desdir=window.showModalDialog(loc);
	//alert(desdir);
	if(desdir!= null){
	  return desdir;
	}
	return "";
}
function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_showHideLayers() { //v3.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
    obj.visibility=v; }
}
function MM_setTextOfLayer(objName,x,newText) { //v3.0
  if ((obj=MM_findObj(objName))!=null) with (obj)
    if (navigator.appName=='Netscape') {document.write(unescape(newText)); document.close();}
    else innerHTML = unescape(newText);
}
function rnfile(filename)
{
	var newname = prompt('��Ҫ���ļ� '+filename+' ����Ϊ:',filename);
	if(newname == null) return "";
	return newname;
}	
-->
</script>
</head>

<body bgcolor="#FFFFFF">
<div align="center">
<%catagory=request("catagory")
  directory="\"
  page=request("page")
  if page="" then
     page=1
  else
     page=cint(page)   
  end if   
if (not session("pass"&catagory)) and not session("admin") then
   Response.Write "<div align=center>�Բ��������ܴ򿪱�ҳ��</div>"
   Response.End
end if
   session("address")="main.asp?page="&page&"&catagory="&catagory
   set conn=server.CreateObject("ADODB.Connection")
   set rs1=server.CreateObject("ADODB.Recordset")
   conn.Open OpenString
   sql="select count(*) from elite where type=0"
   rs1.Open sql,conn,3
   if rs1(0)=0 then
      sql="select * from articles where catagory='"&catagory&"' and score=1"
   else   
      sql="select * from articles where catagory='"&catagory&"' and score=1 and (articles.id not in (select article_id from elite where type=0))"
   end if
   'Response.Write sql
   rs1.close
   rs1.Open sql,conn,3 
   if rs1.EOF then%>
       <div align=center>�Բ���û��δ����ľ�������</div>
       <script language="javascript">
         window.location.replace('view.asp?catagory=<%=catagory%>');
       </script>  
  <%Response.End
   end if
      pgSz=50
      rs1.PageSize=pgSz
      rs1.AbsolutePage=page
      pgCount=rs1.PageCount 
      rsCount=rs1.RecordCount
      %>
  <p><font color="red"><%=catagory%></font>��̳����������</p>
  <p><a href="view.asp?catagory=<%=catagory%>">��������</a>�������½��ļ��С��ƶ��������µȣ�</p>
  <p align="left">û������ľ������£�</p>
  <form method="post" name=frm1 action="deal.asp">
    <table width="100%" border="0">
      <tr align="center" bgcolor="#000066"> 
        <td width="53%"><font color="#CCCCFF">����</font></td>
        <td width="9%"><font color="#CCCCFF">����</font></td>
        <td width="11%"><font color="#CCCCFF">д������</font></td>
        <td width="19%"><font color="#CCCCFF">����</font></td>
      </tr>
      <%
        do while (not rs1.EOF) and ti<pgSz
           ti=ti+1
           author=trim(rs1("author"))
         if not author="" then
                 temp=split(author,"(")
                 username=temp(0)
         end if
         id=rs1("id") %>
      <tr bgcolor="#CCCCFF"> 
        <td width="53%"><a href="..\detail.asp?id=<%=rs1("id")%>&catagory=<%=catagory%>" onclick="return js_callpage(this.href);"><%=trim(rs1("title"))%></a></td>
        <td width="9%"><%=username%></td>
        <td width="11%"><%=trim(rs1("ondate"))%></td>
        <td width="19%"> 
          <input type=hidden name="articleaction_<%=id%>" >
          <a href="javascript:return(false);" onclick="javascript:frm1.articleaction_<%=id%>.value=move('showdir.asp?catagory=<%=catagory%>&time=<%=now%>');if(frm1.articleaction_<%=id%>.value!=''){MM_showHideLayers('div<%=id%>','','show');MM_setTextOfLayer('div<%=id%>','','<font color=gray>�ƶ���:</font>'+articleaction_<%=id%>.value);}">�ƶ�</a>&nbsp;
          <a href="javascript:return(false);" onclick="javascript:frm1.articleaction_<%=id%>.value='';MM_showHideLayers('div<%=id%>','','hide');MM_setTextOfLayer('div<%=id%>','','');">ȡ��</a>
          <div id=div<%=id%> align=center style="visibility: hidden"></div>
        </td>
      </tr>
      <%rs1.MoveNext
        loop%>
    </table>
    <p>
      <input type=hidden name="catagory" value="<%=catagory%>"> 
      <input type="submit" name="mainconfirm" value="ȷ��">
    </p>
  </form>
  <form name="frmPage" method="get" action="main.asp">
    <table width="100%" border="0">
      <tr align="center" bgcolor="#FFCCCC"> 
        <td>��ҳ�� 
          <input type=hidden name="catagory" value=<%=catagory%>>
          <%if page>1 then%> <a href="main.asp?page=<%=page-1%>&catagory=<%=catagory%>">��һҳ</a> 
          &nbsp; <%end if
      if page<pgCount then%> <a href="main.asp?page=<%=page+1%>&catagory=<%=catagory%>">��һҳ</a> 
          &nbsp; <%end if%>������ 
          <select name="page" onchange="if(this.value!='')frmPage.submit();">
            <% for i=1 to pgCount%> 
            <option value=<%=i%><%if i=page then%> selected<%end if%>><%=i%></option>
            <% next %> 
          </select>
          ҳ����<%=rsCount%>ƪ���� </td>
      </tr>
    </table>
    </form>
</div>
</body>
</html>
