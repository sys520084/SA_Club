<!--#include file="../../VAR.asp"-->
<!--#include file="aspFunctions.asp"-->
<%OpenString=DBBS()%>
<html>
<head>
<title>发表文章</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
div,td,p,li,input,textarea{font-size:9pt; line-height:14pt; font-family:宋体;color:black;}
body.mceContentBody {fontsize: 12pt;}
a:link {text-decoration: none; }
a:visited {text-decoration: none;}
a:hover {text-decoration: underline; color: ff0000}
a:active {text-decoration;}
.style1 {color: #FF0000}
-->
</style>
<script language=javascript>
function checkfrm(){
   if(frm1.title.value==""){
     alert("标题不能为空");
     frm1.title.focus();
     return false;
   } 
   <%if request("id")="" then%>
   if(frm1.tags.value==""){
   	 alert("标签Tags不能为空");
     frm1.tags.focus();
     return false;
   }
   <%end if%>
   return true;
   //return confirm("您确认所有信息都填写正确了，并要发表吗？");
} 
function js_callpage(htmlurl) {
var 
newwin=window.open(htmlurl,"homeWin","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=330");
  newwin.focus();
  return false;
}  
</script>
<!-- TinyMCE -->
<script type="text/javascript" src="../tinyfck/tiny_mce.js"></script>

<script type="text/javascript">
var tinyMCEmode = true;
function toogleEditorMode(sEditorID) {
    try {
        if(tinyMCEmode) {
            tinyMCE.removeMCEControl(tinyMCE.getEditorId(sEditorID));
            tinyMCEmode = false;
        } else {
            tinyMCE.addMCEControl(document.getElementById('content'), sEditorID);
            tinyMCEmode = true;
        }
    } catch(e) {
        //error handling
    }
}

tinyMCE.init({
		mode : "textareas",
		theme : "advanced",
		plugins : "table,save,advhr,advimage,advlink,emotions,iespell,insertdatetime,preview,zoom,flash,searchreplace,print,paste,directionality,fullscreen,noneditable,contextmenu",
		theme_advanced_buttons1_add_before : "save,newdocument,separator",
		theme_advanced_buttons1_add : "fontselect,fontsizeselect",
		theme_advanced_buttons2_add : "separator,insertdate,inserttime,preview,zoom,separator,forecolor,backcolor,liststyle",
		theme_advanced_buttons2_add_before: "cut,copy,paste,pastetext,pasteword,separator,search,replace,separator",
		theme_advanced_buttons3_add_before : "tablecontrols,separator",
		theme_advanced_buttons3_add : "emotions,iespell,flash,advhr,separator,print,separator,ltr,rtl,separator,fullscreen",
		theme_advanced_toolbar_location : "top",
		theme_advanced_toolbar_align : "left",
		theme_advanced_statusbar_location : "bottom",
		plugin_insertdate_dateFormat : "%Y-%m-%d",
		plugin_insertdate_timeFormat : "%H:%M:%S",
		extended_valid_elements : "hr[class|width|size|noshade]",
		file_browser_callback : "fileBrowserCallBack",
		paste_use_dialog : false,
		theme_advanced_resizing : true,
		theme_advanced_resize_horizontal : false,
		theme_advanced_link_targets : "_something=My somthing;_something2=My somthing2;_something3=My somthing3;",
		apply_source_formatting : true
	});

	function fileBrowserCallBack(field_name, url, type, win) {
		var connector = "../../filemanager/browser.html?Connector=connectors/php/connector.php";
		var enableAutoTypeSelection = true;

		var cType;
		tinyfck_field = field_name;
		tinyfck = win;

		switch (type) {
			case "image":
				cType = "Image";
				break;
			case "flash":
				cType = "Flash";
				break;
			case "file":
				cType = "File";
				break;
		}

		if (enableAutoTypeSelection && cType) {
			connector += "&Type=" + cType;
		}

		window.open(connector, "tinyfck", "modal,width=600,height=400");
	}
</script>
<!-- /TinyMCE -->
</head>
<body bgcolor="#FBFBF4">
<%atype="bbs"
  if session("user")="" then
		 UrlTail=MyUrlEncode("../forum/new.asp?id=" & request("id") & "&eid="&request("eid"))
         Response.Write "<div align='center'>对不起，您不能发表文章，请<a href='../membership/relogin.asp?UrlTail=" & UrlTail & "'>重新登陆</a>，或者：</div>"
         Response.End
  end if
  set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  isadmin=IsSuperAdmin(Conn,Session("user"))
  
  ThisTail=request("ThisTail")
  UrlTail=request("UrlTail")
  id=request("id")
  eid=request("eid")


  sql="select * from power where username='"&session("user")&"' and power=1"
  'Response.Write sql
  'Response.End 
  set rs=conn.Execute(sql)
  if not rs.eof then
     Response.Write "<div align='center'>对不起，您不能发表文章。因为您的权限被管理员封闭了！</div>"
     Response.End 
  end if   
  rs.close

  if id<>"" then
	id=clng(id)
	atype="bbs"
	 sql="select * from articles where id="&id
	 rs.open sql,conn,3
	 if rs.eof then
		response.write "<div align='center'>对不起，您不能回复此文。可能它已经被删除了！</div>"
		response.end
	 end if
	 catagory=rs("catagory")
	 tags=rs("tags")
     if left(trim(rs("title")),3)<>"回复：" then
        title=server.HtmlEncode("回复："&left(trim(rs("title")),97))
     else
        title=server.HtmlEncode(left(trim(rs("title")),100))
     end if   
     if len(rs("content"))>0 then
       LastContent=GetReplyContent(rs("content"),rs("author"),rs("title"))
     end if
	 LastContent="[请在这里输入你的回复，不要在下面的空间内写，谢谢！劳驾您顺道把我也删了吧]<br><br><br><font color=gray>" & LastContent&"</font><p></p>"
	fatherid=rs("fatherid")
	if fatherid="0" then
		fatherid=id
	end if

  end if
   if eid<>"" then
    sql="select * from articles where id="&eid
	rs.Open sql,conn
	if rs.EOF then
       Response.Write "<div align='center'>对不起，文章已经删除，提取失败！</div>"
       response.End
	end if
	title=rs("title")
	lastcontent=rs("content")
	face=rs("face")
	author=trim(rs("author"))
	catagory=rs("catagory")
	tags=rs("tags")
	fatherid=rs("fatherid")
	'Response.Write admin
	author1=left(author,len(session("user")))
	if rs("score")=1 then
		atype="blog"
	end if
'  author=author1
	if author1<>trim(session("user")) and not isadmin then
		 Response.Write "<div align=center>对不起，您不能更改这篇文章</div>"
		Response.End
	end if
	rs.close
	set rs=nothing
  end if
  saveweb="save.asp"
  if eid<>"" then
	saveweb="deal.asp"
  end if
     %>
<div align="center">发表文章</div>

  <table width="100%" border="1" bgcolor="#EEFFFF" bordercolorlight="#000066" bordercolordark="#CCFFFF" cellpadding="0" cellspacing="0">
  <form method="post" action=<%=saveweb%> name="frm1" onsubmit="return checkfrm();">
    <tr> 
      <td colspan="3" bgcolor="#B09B8D"> <div align="center">题目： 
          <%'Response.Write request("id")
           'Response.Write request("title")
           if title<>"" and id<>"" then%>
			<input type="hidden" name="fatherid" value=<%=fatherid%>>
          <%else%>
			<input type="hidden" name="fatherid" value=0>
          <% 
           end if
		   if eid<>"" then%>
		     <input type="hidden" name="id" value=<%=eid%>>
			 <input type="hidden" name="action" value="edit">
		   <%end if%>
          <input type="hidden" name="name" value="<%=session("user")%>">
		  <%if id<>"" then%>
		  <input type="hidden" name="IsReply" value=1>
		  <%end if%>
		  <input type="hidden" name="ThisTail" value=<%=server.UrlEncode(ThisTail)%>>
		  <input type="hidden" name="UrlTail" value=<%=server.UrlEncode(UrlTail)%>>
		  <input type="hidden" name="pics" value="face01">
          <input type="text" name="title" value="<%=title%>" size="45">
          &nbsp;</div></td>
    </tr>
    <tr>
      <td colspan="3" align="center" bgcolor="#B09B8D"> 
        <%
		'response.write catagory
		if isadmin then
		if id="" then%>
        &nbsp;&nbsp;请选择论坛<select name="catagory">
        <%	sql1="select * from forum_state where fatherid<>0"
			set rs1=conn.Execute(sql1)
			do while not rs1.eof
			%>
			<option value="<%=rs1("catagory")%>" <%if catagory=rs1("catagory") then%>selected<%end if%>><%=rs1("catagory")%></option>
			<%rs1.movenext
			loop
			rs1.close%>			
        </select>
        <%end if
		end if
		%>
</td>
        
    </tr>
	
    
    <tr> 
      <td colspan="3" height="97" align=center> <table align="center" witdh="90%"> 
        <tr><td width="1000">
          <p align="center">内容：<br>
		  <div align="center">
		   <textarea name="content" rows="30" cols="100" style="width: 100%">
		   <%=Server.HTMLEncode(trim(lastcontent))%>
	       </textarea>
		   
		   </div>
		   <div align="center"><!--<a href="#" title="toogle TinyMCE" onclick="toogleEditorMode('content');">HTML源代码</a>-->
		     <input type="reset" name="Submit22" value="HTML源代码" onclick="toogleEditorMode('content');">
		     直接显示源代码</div>
			<div align="center">
			  <p>添加标签（以空格分割）<span class="style1">*</span>                  
			    <%if tags<>"" then
				tags=server.HTMLEncode(tags)
			  end if
			%>
			      <input name="tags" type="text" value="<%=tags%>" size="70">
</p>
			  <p><br>
                <input type="submit" name="save" value="确定并提交">
                <input type="reset" name="Submit2" value="重添">
</p>
			  <hr>
			  <p align="justify"><strong>如何插入图片</strong>？</p>
			  <p align="justify">在输入框的上面的一排按钮中点击“<img src="../tinyfck/themes/advanced/images/image.gif" width="20" height="20">”按钮，可以直接将图片中的URL写到对话框中，或者再次点击右侧的“<img src="../tinyfck/themes/advanced/images/browse.gif" width="20" height="18">”按钮，就会弹出一个对话框，你可以将硬盘上的图片上传到服务器中，然后再选中你上传好的图片，就可以把图片贴到上面的文本框里了。</p>
			  <p align="justify"><strong>如果要上传文件，请操作下面的按钮添加附件：</strong></p>
			  </div>
            </td></tr>
      </table>
	  </td>
    </tr>
	</form>
	<tr><td>
	<table>
<tr>
	<td align="center"><form method="post" name="form1" action="../upload/savedoc.asp?uptype=1" enctype="multipart/form-data" target="_blank">
    添加附件： 
      <input type="file" name="doc">
      <input type="submit" name="confirm" value="确定">
    </form>
	</td>
  </tr>
</table>
	</td></tr>
</table>


</body>
</html>
