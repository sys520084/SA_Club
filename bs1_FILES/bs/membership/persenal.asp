<!--#include file="../../VAR.asp"-->
<!--#include file="../publicmodule/function.asp"-->
<%OpenString=DBBS()%>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="MSHTML 5.00.2920.0" name=GENERATOR>
<META content=none name="Microsoft Border">
<STYLE type=text/css>
DIV {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
}
TD {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
}
A {COLOR: #0099FF; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #FF0000; FONT-FAMILY: "����"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
input,textarea,SELECT {
	FONT-FAMILY: "����";
	FONT-SIZE: 9pt;
	color: #000000;
	background-color: #FFFFFF;
}

.style1 {color: #0000FF}
.style2 {color: #FF0000}
</STYLE>
<script language="javascript">
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
            tinyMCE.addMCEControl(document.getElementById('detailcontent'), sEditorID);
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
<BODY bgColor=#FFFFFF text="#FFFFFF">
</HEAD>
<%if session("user")="" then
         Response.Write "<div align='center'>�Բ��������ܸ��ģ���<a href='relogin.asp' onclick='return js_callpage(this.href)'>���µ�½</a>��</div>"
         Response.End
  end if%>

<%set conn=server.CreateObject("ADODB.Connection")
  conn.Open OpenString
  set rs=server.CreateObject("ADODB.Recordset")
  name=session("user")
  sql="select * from users,detail,memberdes where users.username='"&name&"' and users.username=detail.username and users.username=memberdes.username"
  'response.write sql
  'response.end
  rs.open sql,conn,3
  if rs.eof then
     Response.Write "�Բ���û���ҵ�����û���"
     Response.End
  end if
  nickname=rs("nickname")
  email=rs("email")
  homepage=rs("homepage")
  interests=rs("interething")
  showtitle=rs("showtitle")
  portrait=rs("portrait")
  if isnull(portrait) then
	portrait=""
  end if
         '���㾭��ֵ
      articlecount=cint(rs("articles"))
      cometimes=cint(rs("times"))
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

  rs.close
  sql="select * from memberdes where username='"&name&"'"
  'Response.Write sql
  rs.Open sql,conn
  mytype="new"
  if rs.EOF then
     Response.Write "�Բ���û���ҵ�����û���"
     Response.End
  end if   
  if not rs.EOF then
    readme=trim(rs("readme"))
	detailcontent=rs("detailcontent")
	if isnull(detailcontent) then
		detailcontent=""
	end if
    mytype="old"
  end if  
%>
<DIV align=center> 
 <form action="savedes.asp" method=post name="form1">
    <table width="100%" border="0" cellpadding="2" cellspacing="2" bgcolor="#EFEFEF">
      <tr> 
        <td width="250"><a href="uploadimg.asp?portrait=<%=portrait%>" target="_blank"><img src="<%=ShowPortrait(portrait)%>" alt="����ͷ��" hspace="15" vspace="10" border="0" align="left"  bordercolor="#000000" width=48 height=48></a> 
          <strong>�ǳƣ�</strong><%=nickname%><br> <strong>�����ʼ���</strong><%=email%><br> 
          <strong>����Ȥ���� </strong><%=interests%><br> 
          <%if homepage<>"" then%>
          <strong>������ҳ</strong>��<%=homepage%> <br> 
          <%end if%>
          <strong>����</strong>�� <font color="red"> 
          <%for i=1 to rank%>
          �� 
          <%next%>
          </font></td>
        <td valign="middle"><p align="center"><b>����ҳ�����֣� 
		<script language="JavaScript">   
  function   countlen(){   
      if   (showtitle.value.length>10){   
          alert("�ַ�����������")   
          showtitle.value=showtitle.value.substring(0,10)   
      };
      return true;}
  }
  </script> 
            <input name="showtitle" onKeyUp="countlen()" value="<%=showtitle%>" size="50">
            </b></p>
          <p align="left">���˵����<br>
            <input name="story" type="text" value="<%=readme%>" size="120">
            <% rs.close%>
          </p>
        </td>
      </tr>
    </table>
	<table width="100%" align="center" bgcolor="#EFEFEF"><tr><td align="center"><p>�ڴ˶�����ĸ��Ի�ҳ��:</p>
	      <p class="style1">�����ʵ������д��ʲô����ô�����д����ĸ�����ҳ�����Ǹ��˲��͵���ַ��ע��ֻҪ��ַ���������������ݣ���һ��Ҫ��<span class="style2">http://</span>��ͷ��</p>
	      <p> 
		  <textarea name="detailcontent" rows="30" cols="80" style="width: 100%">
		   <%=Server.HTMLEncode(trim(detailcontent))%>
	       </textarea>
            <input type="hidden" name="detailcontent1" value="">
             <br>
		  <div align="center"><!--<a href="#" title="toogle TinyMCE" onclick="toogleEditorMode('content');">HTMLԴ����</a>-->
		     <input type="reset" name="Submit22" value="HTMLԴ����" onclick="toogleEditorMode('detailcontent');">
		     ֱ����ʾԴ����</div>
          </p>
          <p>
            <input name="confirm" type="submit" id="confirm" value="ȷ���޸�">
          </p></td></tr></table>
  </form>
  <p><BR>
  </p>
</DIV>
</BODY></HTML>
