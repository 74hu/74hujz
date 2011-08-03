<!-- #include file="h.asp" --><!-- #include file="f.asp" -->
<%Response.clear
Response.ContentType="text/vnd.wap.wml; charset=utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?><!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" & vbnewline
Response.Write "<wml><head><meta http-equiv=""Cache-Control"" content=""no-cache""/><meta http-equiv=""Cache-Control"" content=""max-age=0""/></head>" & vbnewline%>
<card title="网站搜索"><p>
<%dim keyword,sear
keyword=hu(request("keyword"))
if keyword="" then
response.write "网站搜索引擎：<br/>"& chr(13)
response.write "<input emptyok=""true"" name=""keyword"&minute(time)&""&second(time)&""" value=""美女"" title=""请输入关键词""/><br/>"
response.write "搜<anchor>文章<go href=""search.asp"" method=""post""><postfield name=""keyword"" value=""$(keyword"&minute(time)&""&second(time)&")""/><postfield name=""sear"" value=""0""/></go></anchor>"
response.write ".<anchor>标题<go href=""search.asp"" method=""post""><postfield name=""keyword"" value=""$(keyword"&minute(time)&""&second(time)&")""/><postfield name=""sear"" value=""1""/></go></anchor>"
response.write ".<anchor>内容<go href=""search.asp"" method=""post""><postfield name=""keyword"" value=""$(keyword"&minute(time)&""&second(time)&")""/><postfield name=""sear"" value=""2""/></go></anchor><br/>"& chr(13)
response.write "搜<anchor>网页<go href=""http://u.yicha.cn/union/x.jsp"" method=""post""><postfield name=""keyword"" value=""$(keyword"&minute(time)&""&second(time)&")""/><postfield name=""site"" value=""2145930044""/><postfield name=""p"" value=""p""/></go></anchor>"
response.write ".<anchor>图片<go href=""http://u.yicha.cn/union/x.jsp"" method=""post""><postfield name=""keyword"" value=""$(keyword"&minute(time)&""&second(time)&")""/><postfield name=""site"" value=""2145930044""/><postfield name=""p"" value=""i""/></go></anchor>"
response.write ".<anchor>MP3<go href=""http://u.yicha.cn/union/x.jsp"" method=""post""><postfield name=""keyword"" value=""$(keyword"&minute(time)&""&second(time)&")""/><postfield name=""site"" value=""2145930044""/><postfield name=""p"" value=""mp3""/></go></anchor><br/>"& chr(13)
else
sear=request("sear")
if sear="" or IsNumeric(sear)=False then
sear=0
end if
sear=clng(sear)
set rs=Server.CreateObject("ADODB.Recordset")
if sear=1 then
rs.open"select id,title from 74hu_article where title like '%" & keyword & "%' order by id desc",conn,1,1
elseif sear=2 then
rs.open"select id,title from 74hu_article where InStr(1,test,'"&Keyword&"',0)>0 order by id desc",conn,1,1
else
rs.open"select id,title from 74hu_article where InStr(1,test,'"&Keyword&"',0)>0 or title like '%" & keyword & "%' order by id desc",conn,1,1
end if
If Not rs.eof	Then
Dim PageSize,i
PageSize=10
Dim Count,page,pagecount,gopage
gopage="search.asp?keyword="&keyword&"&amp;sear="&sear&"&amp;"
Count=rs.recordcount
page=request("page")
if page="" or isnumeric(page)=false or isnull(page) then page=1
page=int(page)
if page<=0 then page=1
pagecount=(count+pagesize-1)\pagesize
if page>pagecount then page=pagecount
rs.move(pagesize*(page-1))
response.write ("共:"&count&"篇相关文章<br/>")
For i=1 To PageSize
If rs.eof Then Exit For	
Response.write "<a href=""/?aid=art&amp;id="&rs("id")&""">"&i+(page-1)*PageSize&"."&ubb(rs("title"))&"</a><br/>" 
rs.moveNext
Next	
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
if pagecount>1 then response.write "<br/>第"&page&"页 共"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/>页 <a href="""&gopage&"page=$(page)"">跳转</a><br/>"
Else
Response.write "没有符合条件的文章<br/>"
end if
rs.close
set rs=nothing
end if
conn.close
set conn=nothing
response.write "<br/><a href='/?aid=guest'>&gt;&gt;给我们提意见</a><br/><a href='/?aid=index'>"&waptitle&"</a>-<a href='/?aid=map'>导航</a>-<a href='/?aid=shuqian'>收藏</a><br/>"
response.write""&wapbei&"</p></card></wml>"
response.end%>