<!--#include file="../VAR.asp"-->
<!--#include file="publicmodule/function.asp"-->
<!--#include file="forum/aspFunctions.asp"-->
<% 
OpenString=DBBS()
xmlfile=server.mappath("feed.xml") 
Set fso = CreateObject("Scripting.FileSystemObject") 
Set MyFile = fso.CreateTextFile(xmlfile,True) 
MyFile.WriteLine("<?xml version=""1.0"" encoding=""gb2312""?>") 
MyFile.WriteLine("<rss version=""2.0"" xmlns:content=""http://purl.org/rss/1.0/modules/content/""	xmlns:wfw=""http://wellformedweb.org/CommentAPI/"" xmlns:dc=""http://purl.org/dc/elements/1.1/"">") 
MyFile.WriteLine("<channel>") 
MyFile.WriteLine("<title>集智俱乐部头脑风暴论坛</title>") 
MyFile.WriteLine("<link>http://www.swarmagents.cn</link>")
MyFile.WriteLine("<description>(Swarm Agents Club)</description>")
MyFile.WriteLine("<pubDate>Sun, 25 Apr 2010 14:07:24 +0000</pubDate>")
MyFile.WriteLine("<generator>http://wordpress.org/?v=2.3.2</generator>")
MyFile.WriteLine("<language>en</language>")
set conn=server.CreateObject("ADODB.Connection")
  'Response.Write openstring
  'Response.End 
  conn.Open OpenString
  set rs=server.CreateObject("ADODB.Recordset")
	'sql="select top 10  * from articles where fatherid=0 order by "&w1&"*datediff('d',ondate,"&date&")+"&w2&"*readtimes+"&w3&"*children+"&w4&"*datediff('d',lastdate,"&date&")"
	sql="select top 10 * from articles where fatherid=0 order by lastdate desc"
	'Response.Write sql
	'Response.End 
	rs.Open sql,conn,3
	if rs.EOF then
		response.end
	else	
  total=0
	do while not rs.EOF 
         id=rs("id")
		 ondate=rs("ondate")&"  "&rs("ontime")
MyFile.WriteLine("<item>")
MyFile.WriteLine("<title>"&rs("title")&"</title>")
MyFile.WriteLine("<link>http://www.swarmagents.cn/bs/viewforum.asp?id="&id&"</link>")
MyFile.WriteLine("<pubDate>"&FormatDateTime(rs("lastdate"),1)&"</pubDate>")
MyFile.WriteLine("<description><![CDATA[")
MyFile.WriteLine(GetAbstract(rs("content"),300))
MyFile.WriteLine("]]></description>")
MyFile.WriteLine("</item>")
		rs.MoveNext
    loop
	end if
	rs.close
MyFile.WriteLine("</channel>")
MyFile.WriteLine("</rss>")
MyFile.Close 
response.redirect("feed.xml")
%> 
