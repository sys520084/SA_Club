<SCRIPT LANGUAGE=VBScript RUNAT=Server>
function Translate(txt)
   result=txt
   select case txt
   case "name"
       result="����"
   case "username"
       result="�û���"
   case "nickname"
       result="�ǳ�"    
   case "sex"
       result="�Ա�"
   case "birthday"
       result="����"
   case "idcard"
       result="����֤����"
   case "province"
       result="ʡ��"
   case "address"
       result="��ַ"
   case "postno"
       result="��������"
   case "tel"
       result="�绰"
   case "education"
       result="ѧ��"
   case "income"
       result="������"
   case "vocation"
       result="ְҵ"
   case "times"
       result="���ʴ���"
   case "signdate"
       result="�Ǽ�ʱ��"
   case "markscore"
       result="��������Ŀ"
   case "height"
       result="����"
   case "weight"
       result="����"
   case "hobby"
       result="����"
   case "school"
       result="��ҵԺУ"
   case "person"
       result="���������"
   case "job"
       result="���빤��"
   case "readme"
       result="���ҽ���"   
   case "homepage"
       result="������ҳ"
   case "interemusic"
       result="����"
   case "interemovie"
       result="ӰƬ"
   case "art"
       result="����"
   case "place"
       result="��ϲ���ĵط�"
   case "lover"
       result="��������"
   case "question"
       result="����"
   case "anwser"
       result="��"
   case "password"
       result="����"
   case "articles"
       result="������"
   case "book"
       result="��ϲ������"
   case "sports"
       result="�����˶�"            
   end select
   
   Translate=result 
end function
function ValueTranslate(txt,table,fields)
  result=txt
  if table="member" then
       select case txt
       case "0"
         if fields="sex" then
           result="��"
         end if  
       case "1"
         if fields="sex" then
           result="Ů"
         end if
       case ""
         result="&nbsp;"
       end select
  end if
  if table="relation" or table="relation,film" then
      select case txt
      case "0"
         if fields="status" then
            result="Ԥ��"
         end if
      case "1"
         if fields="status" then
            result="����"
         end if
      case ""
         result="&nbsp;"
      end select
  end if
  if table="relation" then
     if fields="filmid" then
        sql="select filmname from film where id="&txt
        set recordset=conn.execute(sql)
        if recordset.eof then
          result="��"
        end if 
        result=recordset(0)
        recordset.close
        set recordset=nothing
     end if
  end if   
  if table="detail" then
       select case txt
       case "0"
         if fields="club" then
           result="��"
         end if  
       case "1"
         if fields="club" then
           result="��"
         end if
       case ""
         result="&nbsp;"
       end select
  end if
  if table="detail" and fields="username" then
     result="<a href="&chr(34)&"javascript:sqlval('select users.username as username,users.passwd as password,users.nickname as nickname,users.articles as articles,users.times as times,users.signdate as signdate,users.markscore as markscore,memberdes.height as height,"
     result=result&"memberdes.weight as weight,memberdes.hobby as hobby,memberdes.school as school,memberdes.person as person,memberdes.job as job,memberdes.art as art,"
     result=result&"memberdes.book as book,memberdes.place as place,memberdes.lover as lover,memberdes.sports as sports,memberdes.readme as readme from users,memberdes where users.username=memberdes.username and users.username=\'"&txt&"\'');"&chr(34)&">"&txt&"</a>"
  end if
  if table="users" and fields="username" then
     result="<a href="&chr(34)&"javascript:sqlval('select * from detail where username=\'"&txt&"\'');"&chr(34)&">"&txt&"</a>"
  end if
  if result="" or isnull(result) then
     result="&nbsp;"
  end if 
  ValueTranslate=result
end function
</SCRIPT>