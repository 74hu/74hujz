<!--#include file="h.asp"--><%
response.write "<card title='"&waptitle&"'>" & Chr(13)
response.write "<p align='"&wapconst&"'>" & Chr(13)
if wapfavor="1" then
response.write ""&getfavor&"<br/>" & Chr(13)
end if
if len(countdown) > 7 then
newyear=DateDiff("d",date,Cdate(countdown))
response.write "距"&countname&"还有"&newyear&"天<br/>" & Chr(13)
end if
If Len(waplogo) > 7 Then
response.write "<img src='"&waplogo&"' alt='"&waptitle&"'/><br/>" & Chr(13)
end If
if wapgonggao="1" then
response.write "<a href='?aid=gonggao'><img src='images/msg.gif' alt='.'/>网站发布最新公告!</a><br/>" & Chr(13)
end if

Set rs = Server.CreateObject("adodb.recordset")
rs.open "select lx,class,wmltxt,relid,br,num,classid from 74hu_class where parent=0 order by pid asc", conn, 1, 1
If rs.eOF Then
response.write ("网站建设中..<br/>")
else
rs.Move (0)
j = 1
Do While Not rs.eOF
if rs("lx")="0" then
response.write"<a href='?aid=class&amp;id="&rs("classid")&"'>"&ubb(rs("class"))&"</a>"& chr(13)
elseif rs("lx")="1" then
Response.Write"<a href='?aid=list&amp;id="&rs("relid")&"'>"&ubb(rs("class"))&"</a>"& chr(13)
elseIf rs("lx") = "2" Then
response.write "" & ubbcode(rs("wmltxt")) & "" &Chr(13)
elseIf rs("lx") = "8" Then Call adstr(1)
elseIf rs("lx") = "9" Then
response.write "" & usb(rs("wmltxt")) & "" & Chr(13)
elseIf rs("lx") = "10" Then Call newtitle(rs("num"), rs("relid"))
elseIf rs("lx") = "11" Then Call hottitle(rs("num"), rs("relid"))
elseIf rs("lx") = "12" Then Call wendtitle(rs("num"), rs("relid"))
elseif rs("lx")="19" then
response.write "<input emptyok=""true"" name=""keyword"" value="""" title=""请输入关键词""/><br/>"
response.write "搜<anchor>文章<go href=""search.asp"" method=""post""><postfield name=""keyword"" value=""$(keyword)""/><postfield name=""sear"" value=""0""/></go></anchor>"& chr(13)
response.write "搜<anchor>网页<go href=""http://u.yicha.cn/union/x.jsp"" method=""post""><postfield name=""keyword"" value=""$(keyword)""/><postfield name=""site"" value=""2145930044""/><postfield name=""p"" value=""p""/></go></anchor>"& chr(13)
end If
If rs("br") = "1" Then response.write "<br/>" & Chr(13)
rs.MoveNext
j = j + 1
Loop
end If
rs.close
Set rs = nothing

If waplink = 1 Then

Set Rslc = Server.CreateObject("ADODB.Recordset")
Sqlink = "select ID,namt from 74hu_link Where active =0 order by HU_time desc"
Rslc.open Sqlink, conn, 1, 1
If Rslc.EOF Then
response.write ("暂无首链！<br/>")
End If

aaa = 1
Do While ((Not Rslc.EOF) And aaa <= 8)
response.write "<a href=""?aid=link&amp;id=" & Rslc("id") & "&amp;act=view"">" & ubb(Rslc("NAMT")) & "</a>" & Chr(13)
If aaa Mod 4 = 0 And aaa <> Rslc.RecordCount Then
response.write "<br/>" & Chr(13)
End If
Rslc.MoveNext
aaa = aaa + 1
Loop
Rslc.Close
Set Rslc = Nothing
End If

%>