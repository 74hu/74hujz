<!--#include file="h.asp"-->
<%
Set rs1 = Server.CreateObject("ADODB.Recordset")
sql1="Select class from 74hu_class where classid="&id
rs1.open sql1,conn,1,1 
if rs1.eof then
response.redirect "?aid=index"
response.end
end if
classname=rs1("class")

rs1.close
set rs1=nothing

Response.Write"<card title='"&classname&"-"&waptitle&"'><p>"

set rs = server.createobject("adodb.recordset")
rs.open"select lx,class,classid,wmltxt,num,relid,br from 74hu_class where parent="&id&" order by pid asc",conn,1,1
if rs.eof then 
response.write("栏目建设中..<br/>")
else
rs.Move(0)
j=1
do while not rs.EOF 
if rs("lx")="0" then
response.write"<a href='?aid=class&amp;id="&rs("classid")&"'>"&ubb(rs("class"))&"</a>"& chr(13)
elseif rs("lx")="1" then
Response.Write"<a href='?aid=list&amp;id="&rs("relid")&"'>"&ubb(rs("class"))&"</a>"& chr(13)
elseif rs("lx")="2" then
Response.Write""&ubbcode(rs("wmltxt"))&""& chr(13)
ElseIf rs("lx")="8" Then Call adstr(1)
elseif rs("lx")="9" then
Response.Write""&usb(rs("wmltxt"))&""& chr(13)
elseif rs("lx")="10" then call newtitle(rs("num"),rs("relid"))
elseif rs("lx")="11" then call hottitle(rs("num"),rs("relid"))
elseif rs("lx")="12" then call wendtitle(rs("num"),rs("relid"))
elseif rs("lx")="19" then
response.write "<input emptyok=""true"" name=""keyword"" value="""" title=""请输入关键词""/><br/>"
response.write "搜<anchor>文章<go href=""search.asp"" method=""post""><postfield name=""keyword"" value=""$(keyword)""/><postfield name=""sear"" value=""1""/></go></anchor>"& chr(13)
response.write "搜<anchor>网页<go href=""http://u.yicha.cn/union/x.jsp"" method=""post""><postfield name=""keyword"" value=""$(keyword)""/><postfield name=""site"" value=""2145930044""/><postfield name=""p"" value=""p""/></go></anchor>"& chr(13)
end if
if rs("br")="1" then response.write "<br/>"
rs.MoveNext
j=j+1
loop
end if
rs.close
set rs=nothing

Response.write "-----------<br/>"
if adsetkf("ads3")=1 then
call adstrs(3,2)
end if
%>