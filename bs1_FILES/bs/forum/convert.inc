<SCRIPT LANGUAGE=VBScript RUNAT=Server>
function convert(str)
   if str="" then
      exit function
   end if
   select case trim(str)
   case "环球影视"
      out="0movie"
   case "摇滚音乐"
      out="0rock"
   case "发现之旅"         
      out="0discovery"
   case "流行音乐"
      out="0pop"
   case "民族音乐"
      out="0folk"
   case "国外流行乐"
      out="0seapop"
   case "西方古典乐"
      out="0classical"
   case "音乐软件"
      out="0software"
   case "音乐硬件"
      out="0hardware"
   case "CD发烧友"
      out="0hifi"
   case "乐器发烧友"
      out="0instrument"
   case "吉他"
      out="0guitar"
   case "意见与建议"
      out="0comment"
   case "电子音乐"
      out="0electron"   
   case "另类艺术"
      out="0otherart"   
   case else
      out=""
   end select
   if out<>"" then
      out=out&".gif"
   end if
   convert=out                                 
end function
</script>