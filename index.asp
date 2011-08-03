<!--#include file="f.asp"--><%
aid=request.QueryString("aid")
if aid="" or isnull(aid) then
aid="index"
end if

id=request.QueryString("id")
if id="" or isnumeric(id)=false or isnull(id) then
id=1
end if

Response.Buffer=True
Response.CacheControl="no-cache" 
Response.clear
Response.ContentType="text/vnd.wap.wml; charset=utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?><!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" & vbnewline
Response.Write "<wml><head><meta http-equiv=""Cache-Control"" content=""no-cache""/><meta http-equiv=""Cache-Control"" content=""max-age=0""/></head>" & vbnewline

if aid = "index" then%>
<!--#include file="i.asp"-->
<%elseif aid = "class" then%>
<!--#include file="c.asp"-->
<%elseif aid = "list" then%>
<!--#include file="l.asp"-->
<%elseif aid = "art" then%>
<!--#include file="a.asp"-->
<%elseif aid = "dis" then%>
<!--#include file="d.asp"-->
<%elseif aid = "diss" then%>
<!--#include file="ds.asp"-->
<%elseif aid = "guest" then%>
<!--#include file="g.asp"-->
<%elseif aid = "link" then%>
<!--#include file="y.asp"-->
<%elseif aid = "url" then%>
<!--#include file="u.asp"-->
<%elseif aid = "shuqian" then%>
<!--#include file="sq.asp"-->
<%elseif aid = "map" then%>
<!--#include file="m.asp"-->
<%elseif aid = "gonggao" then%>
<!--#include file="gg.asp"-->
<%else%><!--#include file="i.asp"-->
<!--#include file="e.asp"-->