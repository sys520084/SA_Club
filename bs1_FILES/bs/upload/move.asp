<!--#include file="..\..\var.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<!--#include file="../forum/aspFunctions.asp"-->

<%OpenString=DBBS()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<style>
<!--
div,td,p,li{font-size:9pt; line-height:14pt; font-family:??;color:white;}
a:link {text-decoration: none; color: yellow }
a:visited {
	text-decoration: none;
	color: #FFFF00;
}
a:hover {
	text-decoration: underline;
	color: #FF9900;
}
a:active {text-decoration;
	color: #FF9900;
}
.unnamed1 {  font-family: "����"; font-size: 9pt; font-style: normal; line-height: 100%; font-weight: normal; color: #99FFFF}
.unnamed2 {  font-family: "����"; font-size: 9pt; line-height: 100%; font-weight: normal; color: #294C39}
-->
</style>
<script language="JavaScript">
function move(loc)
{
	var desdir=window.showModalDialog(loc);
	//alert(desdir);
	if(desdir!= null){
	  return desdir;
	}
	return "";
}
function check(){
	var len=form1.catalog.length;
    for(var i=len-1;i>=0;i--)
    {        
       form1.Tcatalog.value+=form1.catalog.options[i].value+",";
    }
	//alert(form1.Tcatalog.value);
	return true
}
function addCatalog(){
	var s=new Array();
	id=move("../../thesis/admin/showdir.asp?time=<%=time()%>");
	if(id!=""){
		s=id.split(",");
		//String[] s=new String[2];
		//s=id.split(",");
		form1.catalog.options.add(new Option(s[0],s[1]));
	}
}
function deleteItem(){
	var len=form1.catalog.length;
    for(var i=len-1;i>=0;i--)
    {        
        if(form1.catalog.options[i].selected)
        {
            form1.catalog.remove(i);
        }
    }
}
</script>
</head>

<body bgcolor="#043F80" link="#00FFFF" alink="#00FFFF" leftmargin="0" topmargin="0">
<div align=center>
<%

  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  set rs=server.CreateObject("ADODB.Recordset")
  
  isadmin=IsSuperAdmin(Conn,Session("user"))
  if not isadmin then
         Response.Write "<div align='center'>�Բ���������ʹ����ҳ</div>"
         Response.End
  end if
  id=request("id")
    sql="select * from files where id="&id
	rs.Open sql,conn,3
	if rs.EOF then
		Response.Write "<div align='center'>�Բ������ݲ�����������û�ҵ�ָ���ļ�¼</div>"
		Response.End
	end if
	file_name=rs("filename")
	fileExt=rs("extendname")
	user=rs("username")
	filesize=rs("bytes")
	title=rs("title")
	des=rs("des")
	author=rs("author")
	journal=rs("journal")
	publishtime=rs("publishtime")
	keywords=rs("keywords")
	rs.Close
%>
    <form name="form1" method="post" action="savemove.asp" onSubmit="javascript:check();">
      <table width="100%" border="1" cellpadding="1" cellspacing="1" bordercolor="#000033" bgcolor="#6666FF">
        <tr> 
          <td bgcolor="#000033">�ļ�����</td>
          <td><%=file_name%><input name="id" type="hidden" value=<%=id%> id="id"></td>
        </tr>
        <tr> 
          <td bgcolor="#000033">��չ����</td>
          <td><%=fileExt%> <input name="extname" type="hidden" value=<%=fileExt%> id="extname2"></td>
        </tr>
        <tr> 
          <td bgcolor="#000033">�ϴ��û���</td>
          <td><%=user%><input name="user" type="hidden" value=<%=user%>></td>
        </tr>
        <tr> 
          <td bgcolor="#000033">�ļ���С��</td>
          <td><%=filesize%> <input name="bytes" type="hidden" value=<%=filesize%> id="bytes"></td>
        </tr>
        <tr>
          <td bgcolor="#000033">���⣺</td>
          <td><input name="title" type="text" id="title" value="<%=title%>" size="50"></td>
        </tr>
        <tr> 
          <td bgcolor="#000033">����ժҪ��</td>
          <td><textarea name="abstract" cols="60" rows="4" id="textarea"><%=des%></textarea></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">�ļ�������</td>
          <td><textarea name="description" cols="60" rows="4" id="textarea"><%=des%></textarea></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">�����ƪ������ѧ�����ģ���ô���ṩ������Ϣ��</td>
          <td></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">�ؼ��֣�</td>
          <td><input name="keywords" id="textarea" value="<%=keywords%>"></td>
        </tr><tr> 
          <td bgcolor="#000033">�������ߣ�</td>
          <td><textarea name="author" type="text" cols="60" rows="4" id="author"><%=author%></textarea></td>
        </tr>
        <tr>
          <td bgcolor="#000033">����ʱ�䣺</td>
          <td><input name="publishtime" type="text" id="publishtime" value="<%=publishtime%>"></td>
        </tr>
		<tr> 
          <td bgcolor="#000033">������/�ļ���</td>
          <td><input name="journal" type="text" id="journal" value="<%=journal%>"></td>
        </tr>
		<tr>
		  <td>����/����</td>
		  <td><label>
		    <input name="mytype" type="radio" value="article" checked>
		    ����</label>
	        <label>
	        <input type="radio" name="mytype" value="program">
	        ����</label></td>
	    </tr>
		<tr>
          <td>���<font color="#FF0000">*</font>��</td>
		  <td><input name="choose" type="button" id="choose" value="��ѡ�����" onClick="javascript:addCatalog();">
              <select name="catalog"  style="WIDTH: 12em" size="5" multiple id="catalog">
              </select>
              <input type="hidden" name="type" value=<%=asktype%>>
              <input name="delete" type="button" value="ɾ��" onClick="javascript:deleteItem()">
              <input name="Tcatalog" type="hidden">
              <font color=red>ע����ѡ��ߵİ�ťѡ�����</font></td>
	    </tr>
		<tr>
          <td>Դ�ļ�λ��<font color="#FF0000">*</font>��</td>
		  <td><div align="left">
              <textarea name="filename" cols="60" rows="4"><%
			  if left(file_name,len("http://"))="http://" then
			  	response.write file_name
				else
				if instr(file_name,"_")<>0 then
					response.write left(file_name,instr(file_name,"_")-1)&"."&fileExt
				else
					response.write file_name
				end if
			  end if
			  %></textarea>
          </div></td>
	    </tr>
      </table>
      <p> 
        <input type="submit" name="Submit" value="�����ļ����������ݿ�">
        <input name="re" type="reset" id="re" value="��д">
		<input type="hidden" name="uptype" value=<%response.write Request.QueryString("uptype")%>>
      </p>
      <p>&nbsp; </p>
    </form>

</div>
</body>
</html>
