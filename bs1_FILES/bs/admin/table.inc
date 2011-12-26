<SCRIPT LANGUAGE=VBScript RUNAT=Server>
function Translate(txt)
   result=txt
   select case txt
   case "name"
       result="姓名"
   case "username"
       result="用户名"
   case "nickname"
       result="昵称"    
   case "sex"
       result="性别"
   case "birthday"
       result="生日"
   case "idcard"
       result="身份证号码"
   case "province"
       result="省份"
   case "address"
       result="地址"
   case "postno"
       result="邮政编码"
   case "tel"
       result="电话"
   case "education"
       result="学历"
   case "income"
       result="月收入"
   case "vocation"
       result="职业"
   case "times"
       result="访问次数"
   case "signdate"
       result="登记时间"
   case "markscore"
       result="好文章数目"
   case "height"
       result="身高"
   case "weight"
       result="体重"
   case "hobby"
       result="爱好"
   case "school"
       result="毕业院校"
   case "person"
       result="最佩服的人"
   case "job"
       result="理想工作"
   case "readme"
       result="自我介绍"   
   case "homepage"
       result="个人主页"
   case "interemusic"
       result="音乐"
   case "interemovie"
       result="影片"
   case "art"
       result="艺人"
   case "place"
       result="最喜欢的地方"
   case "lover"
       result="梦中情人"
   case "question"
       result="问题"
   case "anwser"
       result="答案"
   case "password"
       result="密码"
   case "articles"
       result="文章数"
   case "book"
       result="最喜欢的书"
   case "sports"
       result="体育运动"            
   end select
   
   Translate=result 
end function
function ValueTranslate(txt,table,fields)
  result=txt
  if table="member" then
       select case txt
       case "0"
         if fields="sex" then
           result="男"
         end if  
       case "1"
         if fields="sex" then
           result="女"
         end if
       case ""
         result="&nbsp;"
       end select
  end if
  if table="relation" or table="relation,film" then
      select case txt
      case "0"
         if fields="status" then
            result="预定"
         end if
      case "1"
         if fields="status" then
            result="结算"
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
          result="无"
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
           result="是"
         end if  
       case "1"
         if fields="club" then
           result="否"
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