<SCRIPT LANGUAGE=VBScript RUNAT=Server>
function convert(str)
   if str="" then
      exit function
   end if
   select case trim(str)
   case "����Ӱ��"
      out="0movie"
   case "ҡ������"
      out="0rock"
   case "����֮��"         
      out="0discovery"
   case "��������"
      out="0pop"
   case "��������"
      out="0folk"
   case "����������"
      out="0seapop"
   case "�����ŵ���"
      out="0classical"
   case "��������"
      out="0software"
   case "����Ӳ��"
      out="0hardware"
   case "CD������"
      out="0hifi"
   case "����������"
      out="0instrument"
   case "����"
      out="0guitar"
   case "����뽨��"
      out="0comment"
   case "��������"
      out="0electron"   
   case "��������"
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