 
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
   
   if(frm1.tags.value==""){
   	 alert("标签Tags不能为空");
     frm1.tags.focus();
     return false;
   }
   
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
 
<div align="center">发表文章</div>
 
  <table width="100%" border="1" bgcolor="#EEFFFF" bordercolorlight="#000066" bordercolordark="#CCFFFF" cellpadding="0" cellspacing="0">
  <form method="post" action=save.asp name="frm1" onsubmit="return checkfrm();">
    <tr> 
      <td colspan="3" bgcolor="#B09B8D"> <div align="center">题目： 
          
			<input type="hidden" name="fatherid" value=0>
          
          <input type="hidden" name="name" value="迷糊">
		  

		  <input type="hidden" name="ThisTail" value=>
		  <input type="hidden" name="UrlTail" value=>
		  <input type="hidden" name="pics" value="face01">
          <input type="text" name="title" value="" size="45">
          &nbsp;</div></td>
    </tr>
    <tr>
      <td colspan="3" align="center" bgcolor="#B09B8D"> 
        
</td>
        
    </tr>
	
    
    <tr> 
      <td colspan="3" height="97" align=center> <table align="center" witdh="90%"> 
        <tr><td width="1000">
          <p align="center">内容：<br>
		  <div align="center">
		   <textarea name="content" rows="30" cols="100" style="width: 100%">
		   
	       </textarea>
		   
		   </div>
		   <div align="center"><!--<a href="#" title="toogle TinyMCE" onclick="toogleEditorMode('content');">HTML源代码</a>-->
		     <input type="reset" name="Submit22" value="HTML源代码" onclick="toogleEditorMode('content');">
		     直接显示源代码</div>
			<div align="center">
			  <p>添加标签（以空格分割）<span class="style1">*</span>                  
			    
			      <input name="tags" type="text" value="" size="70">
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

