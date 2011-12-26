<SCRIPT LANGUAGE=VBScript RUNAT=Server>
function MyhtmlEncode(Str)
	if isnull(str) then
		str=""
	end if
    dim result
    dim l
    l=len(str)
    result=""
	dim i
	for i = 1 to l
	    select case mid(str,i,1)
	           case "<"
	               'Response.Write "<font color='red'>"&result&"bf</font>"
	                result=result+"&lt;"
	                'Response.Write "<font color='red'>"&result&"</font>"
	           case ">"
	                result=result+"&gt;"
	           case chr(34)
	                result=result+"&quot;"
	           case "&"
	              If lcase(Mid(Str, i, 6)) <> "&nbsp;" and lcase(Mid(Str,i,6))<>"&quot;" and lcase(mid(str,i,4))<>"&lt;" and lcase(mid(str,i,4))<>"&gt;" and lcase(mid(str,i,5))<>"&amp;" Then
                    result = result + "&amp;"
                  else
                    result=result+"&"
                  End If
	           case chr(9)
	                result=result+"&nbsp;&nbsp;&nbsp;&nbsp;"
	           case chr(32)	           
	                'result=result+"&nbsp;"
	                if i+1<=l and i-1>0 then
	                   if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then	                      
	                      result=result+"&nbsp;"
	                   else
	                      result=result+" "
	                   end if
	                else
	                   result=result+"&nbsp;"	                    
	                end if
	           'case "'"
	               'result=result+"''"
	           case chr(13)
	                result=result&"<br>"    
	           case else
	                result=result+mid(str,i,1)
	     end select
       next
       MyhtmlEncode= result
end function                           
function ShowPortrait(portrait)
	
	if isnull(portrait) then
		portrait=""
	end if
	
	if len(portrait)>4 then
		if left(portrait,4)="pre\" then
			portrait=portrait
		else
			'portrait=server.URLEncode(left(portrait,len(portrait)-4))&right(portrait,4)
		end if
		
	end if
	address=VirtualAddress()&"/complexeconomics/bs/userface/"&portrait
	if len(portrait)<1 then
		address=VirtualAddress()&"/complexeconomics/bs/userface/ci.gif"
	end if
	ShowPortrait=address
end function
function SqlTran(Str)
    dim result
    dim l
    l=len(str)
    result=""
	dim i
	for i = 1 to l
	    select case mid(str,i,1)
	           case "'"
	               result=result+"''"
	           case else
	               result=result+mid(str,i,1)
	     end select
       next
       SqlTran= result
end function
function getFirstImage(content)
	set regex = new regexp
    regex.ignorecase = true
    regex.global = true
		
	regex.pattern="<img(.*)src\s?\=\s?("")([^"">]+)"""
	'regex.pattern = "<img(.*?)src\s?\=\s?(\u0022?)([^\u0022\/>]+)"
	set matches = regex.execute(content)
	img_address=""
    if regex.test(content) then
		'response.write matches(0)
		img_address = matches(0).submatches(2)
	end if
	getFirstImage=img_address
end function

function Entitle(value)
  dim out
  out=""
if value>=0 then
   if value<=100 then
      out="布衣"
   end if
   if value>100 and value<=200 then
      out="士兵"
   end if
   if value>200 and value<=400 then
      out="偏将"
   end if
   if value>400 and value<=500 then
      out="刺史"
   end if         
   if value>500 and value<=600 then
      out="州牧"
   end if 
   if value>600 and value<=700 then
      out="中郎将"
   end if
   if value>700 and value<=800 then
      out="执金吾"
   end if  
   if value>800 and value<=900 then
      out="光禄勋"
   end if  
   if value>900 and value<=1000 then
      out="男爵"
   end if  
   if value>1000 and value<=2000 then
      out="子爵"
   end if  
   if value>2000 and value<=3000 then
      out="伯爵"
   end if  
   if value>3000 and value<=4000 then
      out="侯爵"
   end if  
   if value>4000 and value<=5000 then
      out="公爵"
   end if  
   if value>5000 and value<=6000 then
      out="司空"
   end if  
   if value>6000 and value<=7000 then
      out="司徒"
   end if
   if value>7000 and value<=8000 then
      out="大司马"
   end if
   if value>8000 and value<=10000 then
      out="大将军"
   end if
   if value>10000 then
      out="霸王"
   end if
end if
Entitle=out
end function      
</SCRIPT>