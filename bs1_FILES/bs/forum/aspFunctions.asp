<%
    Dim MarkCount
    MarkCount = 55
    Dim Marks()
    ReDim Marks(MarkCount)
    Marks(1) = "<img "
    Marks(2) = "<a "
    Marks(3) = "</a>"
    Marks(4) = "<sub>"
    Marks(5) = "</sub>"
    Marks(6) = "<sup>"
    Marks(7) = "</sup>"
	Marks(8) = "<P"
	Marks(9) = "</P>"
	Marks(10) = "<BR"
	Marks(11)="<strong>"
	Marks(12)="</strong>"
	Marks(13)="<u>"
	Marks(14)="</u>"
	Marks(15)="<em>"
	Marks(16)="</em>"
	Marks(17)="<OL>"
	Marks(18)="</OL>"
	Marks(19)="<Li>"
	Marks(20)="</li>"
	Marks(21)="<UL>"
	Marks(22)="</UL>"
	Marks(23)="<SPAN "
	Marks(24)="</SPAN>"
	Marks(25)="<font"
	Marks(26)="</font>"
	Marks(27)="<div"
	Marks(28)="</div>"
	Marks(29)="<hr"
	Marks(30)="</hr>"
	Marks(31)="<b>"
	Marks(32)="</b>"
	Marks(33)="<I"
	Marks(34)="</I>"
	Marks(35)="<h2>"
	Marks(36)="<h3>"
	Marks(37)="</h3>"
	Marks(38)="</h2>"
	Marks(39)="<h1 "
	Marks(40)="</h1>"
	Marks(41)="<h2 "
	Marks(42)="<h3 "
	Marks(43)="<h1>"
	Marks(44)="<table"
	Marks(45)="</table>"
	Marks(46)="<tbody"
	Marks(47)="</tbody>"
	Marks(48)="<tr"
	Marks(49)="</tr>"
	Marks(50)="<td"
	Marks(51)="</td>"
	Marks(52)="<h4>"
	Marks(53)="</h4>"
	Marks(54)="<span>"
	Marks(55)="</span>"
Function IsSuperAdmin(conn,name)
	dim rsf
	dim sql
	'response.write name
	'response.end
	if isnull(name) or name="" then
		IsSuperAdmin=False
		exit function
	end if
	set rsf=server.CreateObject("ADODB.Recordset")
	sql="select * from power where username='"&name&"' and (power=0 or power=2)" 
	'response.write sql
	'response.end
	rsf.Open sql,conn
	if not rsf.eof then
		IsSuperAdmin=True
		Exit Function
	end if
	rsf.Close
  IsSuperAdmin=False
End Function
Function MyHtmlEncode(Str)
    dim result
    result=replace(str,"'","''")
    MyHtmlEncode= result
End function
Function ViewImg(str)
    Dim result
    Found = False
	str=replace(str,"[upload]","<a href='")
	str=replace(str,"[/mid]","'>")
	str=replace(str,"[/upload]","</a>")

     For iti = 1 To Len(str)
        Char = Mid(str, iti, 1)
        Select Case Char
        Case "<"
            For jtj = 1 To MarkCount
                If lcase(Mid(str, iti, Len(Marks(jtj)))) = lcase(Marks(jtj)) Then
                    Found = True
                End If
            Next
            If Found Then
                result = result & Char
            Else
                result = result & "&lt;"
            End If
        Case ">"
            If Found Then
                result = result & Char
                Found = False
            Else
                result = result & "&gt;"
            End If
        Case " "
			if not Found and Mid(str,iti+1,1)=" " then
				result=result & "&nbsp;"
			else
				result = result & Char
			end if
	case chr(13)
		if iti>1 then
			temp=Mid(str,iti-1,1)
		end if
			if not found and iti>1 and not temp=">" then
			  result=result & "<br>"
			else
				result = result & Char
			end if
        Case Else
            result = result & Char
        End Select
    Next
    ViewImg = result
End Function

Function GetAbstract(str1,num)
  Dim regEx,strOut
  Set regEx = New RegExp
  regEx.Pattern = "\n"
  regEx.IgnoreCase = True
  regEx.Global = True 
  strOut = regEx.Replace(str1, " ")
  Set regEx1 = New RegExp
  regEx1.Pattern = "<.*?>"
  regEx1.IgnoreCase = True 
  regEx1.Global = True 
  strOut = regEx1.Replace(strOut, "")
  strOut=replace(strOut,"&nbsp;"," ")
  strOut=trim(strOut)

	max=num
	if max>Len(strOut) then
		max=Len(strOut)
	end if
	strOut=Left(strOut,max)
	GetAbstract=strOut
end Function

Function GetAbstract1(str1,num)
		state=0
		Found=false
		substr=""
		strOut=str1
		max=num
		if max>Len(str1) then
			max=Len(str1)
		end if
		strOut=Left(strOut,max)
		For jtj = 1 To MarkCount
			strOut=replace(lcase(strOut),lcase(Marks(jtj)),"")
		next
		GetAbstract=Replace(Replace(strOut,"<",""),">","")
'response.write Marks(2)
	'For iti = 1 To max
    '    Char = Mid(str1, iti, 1)
    '    Select Case Char
    '    Case "<"
    '        For jtj = 1 To MarkCount
    '            If lcase(Mid(str, iti, Len(Marks(jtj)))) = lcase(Marks(jtj)) Then
    '                Found = True
	'				state=1
    '            End If
    '        Next
	'	Case ">"
	'		strOut=replace(strOut,substr&">","1111111111111")
	'		'response.write vbCrlf&strOut
	'		substr=""
	'		Found=False
	'	end select
	'	if Found then
	'		substr=substr & Char
	'	end if
	'next
	'GetAbstract=strOut
end Function
Function SearchingTopics(SearchTable)
  dim catagory,title,author,fromDate,toDate,content,order,where,sql,urlTail,tags
  catagory=MyHtmlEncode(request("catagory"))
  subs=MyHtmlEncode(request("subs"))
  title=MyHtmlEncode(request("title"))
  author=MyHtmlEncode(request("author"))
  authortype=MyHtmlEncode(request("authortype"))
  fromDate=MyHtmlEncode(request("fromdate"))
  toDate=MyHtmlEncode(request("todate"))
  content=MyHtmlEncode(request("content"))
  order=MyHtmlEncode(request("order"))
  tags=MyHtmlEncode(request("tags"))
  score=MyHtmlEncode(request("score"))
  if request("shortsearch")="ok" then
  	searchValue=MyHtmlEncode(request("searchvalue"))
	if request("item")="title" then
		title=searchValue
	end if
	if request("item")="author" then
		author=searchValue
		authortype=1
	end if
	if request("item")="tags" then
		tags=searchValue
	end if
	if request("item")="content" then
		content=searchValue
	end if
	subs="no"
  end if
  where=""
  sql=""
  urlTail=""
  if subs="no" or subs="" then
	where= " fatherid=0 and"
  else
	urlTail=urlTail & "subs=yes&"
  end if
  if catagory<>"" then
	where = where & " catagory='" & catagory & "' and"
	urlTail=urlTail & "catagory=" & server.UrlEncode(catagory) & "&"
  end if
  if title<>"" then
	where = where & " title like '%" & title & "%' and"
	urlTail=urlTail & "title=" & server.UrlEncode(title) & "&"
  end if
  if author<>"" then
	if authortype="1" then
		where=where & " author like '%" & author & "%' and"
	else
		authortype="0"
		where=where & " (author='" & author &"' or author like '" & author & "(%') and"
	end if
	' or author like '%("& author &")' or author='" & author & "') and"
	urlTail=urlTail & "author=" & server.UrlEncode(author) & "&authortype=" & server.UrlEncode(authortype) &"&"
  end if
  if tags<>"" then
	where=where & " tags like '%" & tags & "%' and"
	' or author like '%("& author &")' or author='" & author & "') and"
	urlTail=urlTail & "tags=" & server.UrlEncode(tags) & "&"
  end if
  if fromDate<>"" then
	if toDate="" then toDate=date
	where=where & " (ondate between #" & fromDate & "# and #" & toDate & "#) and"
	urlTail=urlTail & "fromdate=" & server.UrlEncode(fromDate) & "&todate=" & toDate & "&"
  end if
  if content<>"" then
	where=where & " [content] like '%" & content & "%' and"
	urlTail=urlTail & "content=" & server.UrlEncode(content) & "&"
  end if
   if score<>"" then
	where=where & " [score] = " & score & " and"
	urlTail=urlTail & "score=" & server.UrlEncode(score) & "&"
  end if
  if order="" then
	order="lastdate desc"
  else
	urlTail=urlTail & "order=" & server.UrlEncode(order) & "&"
  end if
  if right(where,len("and"))="and" then
	where=left(where,len(where)-len("and"))
  end if
  if right(urlTail,len("&"))="&" then
	urlTail=left(urlTail,len(urlTail)-1)
  end if
  sql="select * from " & SearchTable & " where " & where & " order by " & order
  SearchingTopics= sql & vbTab & urlTail
end function
Function GetReplyContent(contentIn,Author,title)
	contemp=">"&Author&"��" & title & "��д����<br>" 
	contemp=contemp&"---------------------------<br>"
	contemp= contemp & left(contentIn,100)&"......"
	GetReplyContent=contemp
end Function
function MyUrlEncode(urlIn)
	dim result
	result=replace(urlIn,"=","*")
	result=replace(result,"&","@")
	MyUrlEncode=server.UrlEncode(result)
end function
function MyUrlDecode(urlIn)
	dim result
	result=replace(urlIn,"*","=")
	result=replace(result,"@","&")
	MyUrlDecode=result
end function
%>