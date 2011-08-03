<!--#include file="h.asp"-->
<%
act=request.QueryString("act")
if act<>"" then

response.write"<card title='站内排行榜'><p>"
if act="top" then
response.write"-<a href='?aid=index'>首页</a>-站内排行Top100<br/>-----------<br/>"
end if
if act="new" then
response.write"-<a href='?aid=index'>首页</a>-站内最新Top100<br/>-----------<br/>"
end if
else
Set rs = Server.CreateObject("ADODB.Recordset")
sql="Select class from 74hu_list where classid="&id
rs.open sql,conn,3,1 
if rs.eof then
rs.close
set rs=Nothing
response.redirect "?aid=index"
response.end
end if
classname=rs("class")
rs.close
set rs=Nothing

response.write"<card title='"&classname&"-"&waptitle&"'><p>"
response.write"-<a href='?aid=index'>首页</a>-"&classname&"-<a href='?aid=list&amp;act=top'>排行</a><br/>-----------<br/>"

end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="Select id,title from 74hu_article where classid="&id&" order by id desc"
rs.open sql,conn,3,1
If Not rs.eof then
if adsetkf("ads1")=1 then
call adstr(1)
response.write"<br/>"
end if
PageSize=listnums
gopage="?aid=list&amp;id="&id&"&amp;"
Count=rs.recordcount
page=request("page")
if Isnull(page) or Isnumeric(page)=Flase then page=1
page=int(page)
if page<=0 or page="" then page=1
pagecount=(count+pagesize-1)\pagesize
if page>pagecount then page=pagecount
rs.move(pagesize*(page-1))
For i=1 To PageSize
If rs.eof Then Exit For

response.write"<a href='?aid=art&amp;id="&rs("id")&"'>"&i+(page-1)*PageSize&"."&ubb(rs("title"))&"</a><br/>"
rs.moveNext
Next
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>&nbsp;"
if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
if pagecount>1 then response.write "(<b>"&page&"</b>/"&pagecount&")"&"<br/><input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/>页 <a href="""&gopage&"page=$(page)"">翻页</a><br/>"

response.write("[相关内容]<br/>")
call wendtitle(4,id)
if adsetkf("ads2")=1 then
call adstr(2)
response.write"<br/>"
end if
Else
response.write"暂时没有文章！<br/>"
end if
rs.close
set rs=nothing%>