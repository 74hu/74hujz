<!--#include file="h.asp"--><%
'
'	七色虎建站系统
'	前台展示文件Search.asp
'	用于展示网站
'	v1.2.4.143a
'	2011.9.3

Dim aid,id,index_asp_head
aid=LCase(getDD("aid","index"))
id=getN("id",1)
cache False'消除缓存

if wapstyle<>"2" then

index_asp_head="<meta http-equiv=""Cache-Control"" content=""no-cache""/><meta http-equiv=""Cache-Control"" content=""max-age=0""/></head>"
getHead index_asp_head,1

Select Case aid
	Case "index" showIndex
	Case "art" showArticle
	Case "list" showList
	Case "link" showLink
	Case "guest" showGuest
	Case "dis" showDiscuss
	Case "class" showClass
	Case "url" showUrl
	Case "diss" showComment
	Case "map" showMap
	Case "gonggao" showReport
	Case "shuqian" showBookmark
	Case Else showIndex
End Select
getEnd "<br/>"&ubbcode(wapbei),1
else
index_asp_head= "<meta http-equiv=""Cache-control"" content=""max-age=0"" />"&_
  "<meta http-equiv=""Cache-control"" content=""no-cache""/>"&_
  "<meta name=""viewport"" content=""width=device-width; initial-scale=1.3;  minimum-scale=1.0; maximum-scale=2.0""/>"&_
  "<meta name=""MobileOptimized"" content=""240""/>"&_
  "<meta name=""format-detection"" content=""telephone=no"" />"& vbnewline &_
  "<style type=""text/css"">"&_
  "body{font-size:14px;width:250px;text-align:center;margin:0 auto;background:#EAEAEA}"&_
  "div{width:245px;text-align:left;word-wrap:break-word;overflow:hidden;background:#FFFFFF}"&_
  ".main{width:240px;border:1px solid #C6C6C6;padding:5px}"&_
  ".nav{width:240px;background:#FFFBE1;border:1px solid #FEBF90}"&_
  ".tle{font-weight:bold;text-align:center}"&_
  "a{text-decoration:none;color:#0A63BB;}"&_
  "a:hover{text-decoration:underline;color:#DE0000;}"&_
  "img,a img{border:none;}"&_
  "form{margin:0px;display: inline;}"&_
  ".tip{color:#DE0000}"&_
  "</style>"& vbnewline
getHead index_asp_head,2

Select Case aid
	Case "index" showsIndex
	Case "art" showsArticle
	Case "list" showsList
	Case "link" showsLink
	Case "guest" showsGuest
	Case "dis" showsDiscuss
	Case "class" showsClass
	Case "url" showsUrl
	Case "diss" showsComment
	Case "map" showsMap
	Case "gonggao" showsReport
	Case "shuqian" showsBookmark
	Case Else showsIndex
End Select
getEnd "<br/>"&ubbcode(wapbei),1
end if

getClose
%>