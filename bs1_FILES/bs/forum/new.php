 
<html>
<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style> 
<!--
div,td,p,li,input,textarea{font-size:9pt; line-height:14pt; font-family:����;color:black;}
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
     alert("���ⲻ��Ϊ��");
     frm1.title.focus();
     return false;
   } 
   
   if(frm1.tags.value==""){
   	 alert("��ǩTags����Ϊ��");
     frm1.tags.focus();
     return false;
   }
   
   return true;
   //return confirm("��ȷ��������Ϣ����д��ȷ�ˣ���Ҫ������");
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
 
<div align="center">��������</div>
 
  <table width="100%" border="1" bgcolor="#EEFFFF" bordercolorlight="#000066" bordercolordark="#CCFFFF" cellpadding="0" cellspacing="0">
  <form method="post" action=save.asp name="frm1" onsubmit="return checkfrm();">
    <tr> 
      <td colspan="3" bgcolor="#B09B8D"> <div align="center">��Ŀ�� 
          
			<input type="hidden" name="fatherid" value=0>
          
          <input type="hidden" name="name" value="�Ժ�">
		  

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
          <p align="center">���ݣ�<br>
		  <div align="center">
		   <textarea name="content" rows="30" cols="100" style="width: 100%">
		   
	       </textarea>
		   
		   </div>
		   <div align="center"><!--<a href="#" title="toogle TinyMCE" onclick="toogleEditorMode('content');">HTMLԴ����</a>-->
		     <input type="reset" name="Submit22" value="HTMLԴ����" onclick="toogleEditorMode('content');">
		     ֱ����ʾԴ����</div>
			<div align="center">
			  <p>��ӱ�ǩ���Կո�ָ<span class="style1">*</span>                  
			    
			      <input name="tags" type="text" value="" size="70">
</p>
			  <p><br>
                <input type="submit" name="save" value="ȷ�����ύ">
                <input type="reset" name="Submit2" value="����">
</p>
			  <hr>
			  <p align="justify"><strong>��β���ͼƬ</strong>��</p>
			  <p align="justify">�������������һ�Ű�ť�е����<img src="../tinyfck/themes/advanced/images/image.gif" width="20" height="20">����ť������ֱ�ӽ�ͼƬ�е�URLд���Ի����У������ٴε���Ҳ�ġ�<img src="../tinyfck/themes/advanced/images/browse.gif" width="20" height="18">����ť���ͻᵯ��һ���Ի�������Խ�Ӳ���ϵ�ͼƬ�ϴ����������У�Ȼ����ѡ�����ϴ��õ�ͼƬ���Ϳ��԰�ͼƬ����������ı������ˡ�</p>
			  <p align="justify"><strong>���Ҫ�ϴ��ļ������������İ�ť��Ӹ�����</strong></p>
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
    ��Ӹ����� 
      <input type="file" name="doc">
      <input type="submit" name="confirm" value="ȷ��">
    </form>
	</td>
  </tr>
</table>
	</td></tr>
</table>
 
 
</body>
</html>

