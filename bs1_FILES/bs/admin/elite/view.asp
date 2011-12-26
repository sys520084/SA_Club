<!--#include file="../../../VAR.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>精华区整理</title>
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
	var newname = prompt('您要将文件 '+filename+' 改名为:',filename);
	if(newname == null) return "";
	return newname;
}	
-->
</script>
</head>

<body bgcolor="#FFFFFF">
<div align="center">
<%catagory=request("catagory")
  id=request("id")
  if id="" then
     id=0
  else
     id=cint(id)
  end if      
  page=request("page")
  if page="" then
     page=1
  end if
     page=cint(page)   
if (not session("pass"&catagory)) and not session("admin") or session("administrator")="" then
   Response.Write "<div align=center>对不起，您不能打开本页！</div>"
   Response.End
end if
   set conn=server.CreateObject("ADODB.Connection")
   conn.Open OpenString
   if id<>0 then
      sql="select * from elite where id="&id
      'Response.Write sql
      'Response.end
      set rs1=conn.Execute(sql)
      directory=trim(rs1("location"))&trim(rs1("title"))&"\"
      if rs1.eof then
         Response.Write "<div align=center>对不起，没有您指定的路径，可能它已经被删除！</div>"
         Response.End
      end if
   else
     directory=request("directory")
     if directory="" then
        directory="\"
     end if   
   end if     
   set rs=server.CreateObject("ADODB.Recordset")
   sql="select * from elite where catagory='"&catagory&"' and location='"&directory&"' order by type desc,id desc"
   'Response.Write sql
   rs.Open sql,conn,3
   if not rs.EOF then
      pgSz=20
      rs.PageSize=pgSz
      rs.AbsolutePage=page
      pgCount=rs.PageCount
      rsCount=rs.RecordCount
   end if
   t=split(directory,"\")
   'Response.Write directory
   'Response.Write t(1)
   for i=1 to ubound(t)-1
       for j=1 to i
          add=add&"\"&t(j)
       next
       add=add&"\"
       out=out&"<a href='view.asp?directory="&server.URLEncode(add)&"&catagory="&catagory&"'>"&t(i)&"</a>\"
       add=""
   next
   out="<a href='view.asp?directory=\&catagory="&catagory&"'>\</a>"&out           
   session("address")="view.asp?directory="&directory&"&page="&page&"&catagory="&catagory
   %>
  <p><font color="red"><%=catagory%></font>论坛精华区整理</p>
  <p><a href="main.asp?catagory=<%=catagory%>">未处理的精华文章</a></p>
  <p align="left">您当前的位置：<%=out%></p>
  <%if rs.EOF then%>
       <div align=center>对不起，该目录下没有精华文章</div>
  <%else%>
  <form method="post" name=frm1 action="deal.asp">
    <table width="100%" border="0">
      <tr align="center" bgcolor="#000066"> 
        <td width="8%"><font color="#CCCCFF">类型</font></td>
        <td width="53%"><font color="#CCCCFF">名称</font></td>
        <td width="9%"><font color="#CCCCFF">整理人</font></td>
        <td width="11%"><font color="#CCCCFF">处理日期</font></td>
        <td width="19%"><font color="#CCCCFF">处理</font></td>
      </tr>
      <%do while not rs.EOF and ti<pgSz
           ti=ti+1%>
      <tr bgcolor="#CCCCCC">
        <td width="8%">
        <%if rs("type")=1 then%>
         目录
        <%else
          showconfirm=true%>
         文章
        <%end if%>
         </td>
        <td width="53%">
        <%if rs("type")=1 then%> 
          <a href="view.asp?id=<%=rs("id")%>&catagory=<%=catagory%>"><%=trim(rs("title"))%></a></td>
        <%else
          article_id=rs("article_id")
          id=rs("id")
          %>
          <a href="..\detail.asp?id=<%=rs("article_id")%>&catagory=<%=catagory%>" onclick="return js_callpage(this.href);"><%=trim(rs("title"))%></a></td>
        <%end if%>
        <td width="9%"><%=trim(rs("deal_man"))%></td>
        <td width="11%"><%=trim(rs("deal_date"))%></td>
        <td width="19%"> 
          <%if rs("type")=1 then%>
          <select name="cataaction" onchange="if(this.value!='')return js_callpage('dealcata.asp?id=<%=rs("id")%>&catagory=<%=catagory%>&action='+this.value);">
            <option value="">操作</option>
            <option value="move">移动</option>
            <option value="rename">重命名</option>
            <option value="del">删除</option>
          </select>
          <%else%>
          <input type=hidden name="articleaction_<%=id%>" >
          <input type=hidden name="delarticle_<%=id%>">
          <input type=hidden name="rename_<%=id%>">
          <a href="javascript:return(false);" onclick="javascript:frm1.delarticle_<%=id%>.value='del';MM_showHideLayers('div<%=id%>','','show');MM_setTextOfLayer('div<%=id%>','','<font color=red>X（删）</font>');frm1.articleaction_<%=id%>.value='';rename_<%=id%>.value='';">删除</a>&nbsp;
          <a href="javascript:return(false);" onclick="javascript:frm1.articleaction_<%=id%>.value=move('showdir.asp?catagory=<%=catagory%>&time=<%=now%>');if(frm1.articleaction_<%=id%>.value!=''){MM_showHideLayers('div<%=id%>','','show');MM_setTextOfLayer('div<%=id%>','','<font color=gray>移动到:</font>'+articleaction_<%=id%>.value);frm1.delarticle_<%=id%>.value='';rename_<%=id%>.value='';}">移动</a>&nbsp;
          <a href="javascript:return(false);" onclick="javascript:frm1.rename_<%=id%>.value=rnfile('<%=replace(trim(rs("title")),"'","\'")%>');if(frm1.rename_<%=id%>.value!=''){MM_showHideLayers('div<%=id%>','','show');MM_setTextOfLayer('div<%=id%>','','<font color=gray>改名为:</font>'+rename_<%=id%>.value);frm1.delarticle_<%=id%>.value='';articleaction_<%=id%>.value='';}">重命名</a>&nbsp;
          <a href="javascript:return(false);" onclick="javascript:frm1.rename_<%=id%>.value='';frm1.articleaction_<%=id%>.value='';frm1.delarticle_<%=id%>.value='';MM_showHideLayers('div<%=id%>','','hide');MM_setTextOfLayer('div<%=id%>','','');">取消</a>
          <div id=div<%=id%> align=center style="visibility: hidden"></div>
          <%end if%>
        </td>
      </tr>
      <%  rs.MoveNext 
        loop
        %> 
     </table>
    <p>
      <input type=hidden name="catagory" value="<%=catagory%>">
      <%if showconfirm then%>
      <input type="submit" name="viewconfirm" value="确定">
      <%end if%>
    </p>
  </form>
  <form name="frmPage" method="get" action="view.asp">
    <table width="100%" border="0">
      <tr align="center" bgcolor="#FFCCCC"> 
        <td>分页： 
          <input type=hidden name="catagory" value="<%=catagory%>">
          <input type=hidden name="directory" value="<%=directory%>">
          <%if page>1 then%> <a href="view.asp?page=<%=page-1%>&catagory=<%=catagory%>&directory=<%=directory%>">上一页</a> 
          &nbsp; <%end if
      if page<pgCount then%> <a href="view.asp?page=<%=page+1%>&catagory=<%=catagory%>&directory=<%=directory%>">下一页</a> 
          &nbsp; <%end if%>跳到： 
          <select name="page" onChange="frmPage.submit();">
            <% for i=1 to pgCount%> 
            <option value=<%=i%><%if i=page then%> selected<%end if%>><%=i%></option>
            <% next %> 
          </select>
          页，共<%=rsCount%>篇精华 </td>
      </tr>
    </table>
    </form>
  <%end if%>
  <form name="frmCreate" action="deal.asp" method="post">
    <table width="100%" border="0">
      <tr align="center" bgcolor="#FFFFCC"> 
        <td>新建文件夹： 
          <input type="text" name="dirname" maxlength="100">
          <input type=hidden name="catagory" value="<%=catagory%>">
          <input type=hidden name="dir" value="<%=directory%>">
          <input type="submit" name="createdir_confirm" value="确定">
          <input type="reset" name="Submit4" value="重写" onClick="window.location.refresh();">
        </td>
      </tr>
    </table>
    </form>
</div>
</body>
</html>
