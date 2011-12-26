<%
const MAX_NUM_MSG=200
const UPFILEPATH="\complexeconomics\bs\files\"
const ALLOW_UP_VALUE=5
const FILE_MAX_SIZE=10485760
const MAX_FULLSPACE=10485760
const DAY_FILES=5
const Domain="http://www.swarmagents.com"
const DomainMain="http://www.swarmagents.cn"
'const Domain="http://www.swarmagents.com"

function WebAddress()
	WebAddress=server.MapPath("/")
end function
function VirtualAddress()
	address=WebAddress()
	'indexa=instr(0,address,server.MapPath("/"))
	if instr(address,server.MapPath("/"))<>0 then
		VirtualAddress=right(address,len(address)-len(server.MapPath("/")))
	else
		VirtualAddress="/clustering"
	end if
end function

function DBBS()
	DbPath=WebAddress &"/complexeconomics/bs/db/database.mdb"
	'Response.Write DbPath
	'Response.End 
	DBBS="DRIVER={Microsoft Access Driver (*.mdb)};DBQ="&DbPath
	'Response.Write DbPath
	'Response.End 
end function
function ExtAddress(i)
	if i=0 then
		ExtAddress="http://www.swarmagents.com"
	end if
end function
%>