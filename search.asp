<!-- #include file="h.asp" --><%
'
'	七色虎建站系统
'	搜索引擎文件Search.asp
'	主要用于文章搜索
'	v1.2.4.143a
'	2011.9.3

Response.clear
Dim search_asp_head
if wapstyle<>"2" then 
search_asp_head="<meta http-equiv=""Cache-Control"" content=""no-cache""/><meta http-equiv=""Cache-Control"" content=""max-age=0""/></head><card title=""网站搜索""><p>"
getHead search_asp_head,1

dim keyword,sear
keyword=getFilter("keyword","")
if keyword="" then
response.write "网站搜索引擎：<br/>"&_
 "<input emptyok=""true"" name=""keyword"&Time_r&""" value=""美女"" title=""请输入关键词""/><br/>"&_
 "搜<anchor>文章<go href=""search.asp"" method=""post""><postfield name=""keyword"" value=""$(keyword"&Time_r&")""/><postfield name=""sear"" value=""0""/></go></anchor>"&_
 ".<anchor>标题<go href=""search.asp"" method=""post""><postfield name=""keyword"" value=""$(keyword"&Time_r&")""/><postfield name=""sear"" value=""1""/></go></anchor>"&_
 ".<anchor>内容<go href=""search.asp"" method=""post""><postfield name=""keyword"" value=""$(keyword"&Time_r&")""/><postfield name=""sear"" value=""2""/></go></anchor><br/>"&_
 "搜<anchor>网页<go href=""http://u.yicha.cn/union/x.jsp"" method=""post""><postfield name=""keyword"" value=""$(keyword"&Time_r&")""/><postfield name=""site"" value=""2145930044""/><postfield name=""p"" value=""p""/></go></anchor>"&_
 ".<anchor>图片<go href=""http://u.yicha.cn/union/x.jsp"" method=""post""><postfield name=""keyword"" value=""$(keyword"&Time_r&")""/><postfield name=""site"" value=""2145930044""/><postfield name=""p"" value=""i""/></go></anchor>"&_
 ".<anchor>MP3<go href=""http://u.yicha.cn/union/x.jsp"" method=""post""><postfield name=""keyword"" value=""$(keyword"&Time_r&")""/><postfield name=""site"" value=""2145930044""/><postfield name=""p"" value=""mp3""/></go></anchor><br/>"
else
sear=getN("sear",0)
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
page=getN("page",0)
if page<=0 then page=1
pagecount=(count+pagesize-1)\pagesize
if page>pagecount then page=pagecount
rs.move(pagesize*(page-1))
response.write ("共:"&count&"篇相关文章<br/>")
For i=1 To PageSize
If rs.eof Then Exit For	
Response.write "<a href=""/?aid=art&amp;id="&rs("id")&""">"&i+(page-1)*PageSize&"."&noubb(rs("title"))&"</a><br/>" 
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
response.write "<br/>"&ubbcode(wapbei)&"</p></card></wml>"
response.end
else

search_asp_head="<meta http-equiv=""Cache-control"" content=""max-age=0"" />"&_
  "<meta http-equiv=""Cache-control"" content=""no-cache""/>"&_
  "<meta name=""viewport"" content=""width=device-width; initial-scale=1.3;  minimum-scale=1.0; maximum-scale=2.0""/>"&_
  "<meta name=""MobileOptimized"" content=""240""/>"&_
  "<meta name=""format-detection"" content=""telephone=no"" />"& vbnewline &_
  "<style type=""text/css"">"&_
  "body{font-size:14px;width:250px;text-align:center;margin:0 auto;background:#EAEAEA}"&_
  "div{text-align:left;background:#FFFFFF}"&_
  ".main{width:240px;border:1px solid #C6C6C6;padding:5px}"&_
  ".nav{width:240px;background:#FFFBE1;border:1px solid #FEBF90}"&_
  "a{text-decoration:none;color:#0A63BB;}"&_
  "a:hover{text-decoration:underline;color:#DE0000;}"&_
  "img,a img{border:none;}"&_
  "form{margin:0px;display: inline;}"&_
  "font{color:#DE0000}"&_
  "</style>"& vbnewline &_
  "<title>网站搜索</title></head><body><div class=""main"">"
getHead search_asp_head,2
keyword=getFilter("keyword","")
if keyword="" then
response.write "<div class=""nav""><a href=""/?aid=index"">首页</a>-网站搜索</div>本站搜索：<br/>"&_
 "<form name=""form1"&Time_r&""" action=""search.asp"" method=""post"">"&_
 "<input name=""keyword"" value="""" type=""text""/><br/>"&_
 "<input type=""radio"" name=""sear"" value=""0"" checked />搜文章<br/>"&_
 "<input type=""radio"" name=""sear"" value=""1"" />搜标题<br/>"&_
 "<input type=""radio"" name=""sear"" value=""2"" />搜内容<br/>"&_
 "<input name=""submit"" value=""开始搜索"" type=""submit""></form><br/>"&_
 "全网搜索：<br/>"&_
 "<form name=""form2"&Time_r&""" action=""http://u.yicha.cn/union/x.jsp"" method=""post"">"&_
 "<input name=""keyword"" value="""" type=""text""/><input name=""site"" value=""2145930044"" type=""hidden""/><br/>"&_
 "<input type=""radio"" name=""p"" value=""p"" checked />搜网页<br/>"&_
 "<input type=""radio"" name=""p"" value=""i"" />搜图片<br/>"&_
 "<input type=""radio"" name=""p"" value=""mp3"" />搜MP3<br/>"&_
 "<input name=""submit"" value=""开始搜索"" type=""submit""></form><br/>"
else
sear=getN("sear",0)
set rs=Server.CreateObject("ADODB.Recordset")
if sear=1 then
rs.open"select id,title from 74hu_article where title like '%" & keyword & "%' order by id desc",conn,1,1
elseif sear=2 then
rs.open"select id,title from 74hu_article where InStr(1,test,'"&Keyword&"',0)>0 order by id desc",conn,1,1
else
rs.open"select id,title from 74hu_article where InStr(1,test,'"&Keyword&"',0)>0 or title like '%" & keyword & "%' order by id desc",conn,1,1
end if
If Not rs.eof	Then
PageSize=10
gopage="search.asp?keyword="&keyword&"&amp;sear="&sear&"&amp;"
Count=rs.recordcount
page=getN("page",0)
if page<=0 then page=1
pagecount=(count+pagesize-1)\pagesize
if page>pagecount then page=pagecount
rs.move(pagesize*(page-1))
response.write ("<div class=""nav"">-<a href=""?aid=index"">首页</a>-<a href=""search.asp"">搜索</a>-搜索结果</div>共:"&count&"篇相关文章<br/>")
For i=1 To PageSize
If rs.eof Then Exit For	
Response.write "<a href=""/?aid=art&amp;id="&rs("id")&""">"&i+(page-1)*PageSize&"."&noubb(rs("title"))&"</a><br/>" 
rs.moveNext
Next	
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
if pagecount>1 then response.write "<br/>第"&page&"页 共"&pagecount&"页<form name=""f"&Time_r&""" action=""search.asp"" method=""post""><input name=""page"" value="""&page&""" maxlength=""2"" size=""3""/>页<input type=""hidden"" name=""keyword"" value="""&keyword&"""/><input type=""hidden"" name=""sear"" value="""&sear&"""/><input type=""submit"" value=""跳转""></form><br/>"
Else
Response.write "没有符合条件的文章<br/>"
end if
rs.close
set rs=nothing
end if
conn.close
set conn=nothing
response.write "<br/>"&ubbcode(wapbei)&"</div></body></html>"
response.end
end if
%>