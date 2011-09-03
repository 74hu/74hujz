<!--#include file="Head.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="mymin.asp"-->
<%
p=trim(request.querystring("p"))
if p="" or IsNumeric(p)=False then p=1

act=trim(request.querystring("act"))
Dim gettxt
call conndata
if act<>"out" then
if p=1 then
HU="今日链入排行"
gettxt="where active=0 and del=0 and DATEDIFF('d', HU_time, date())=0"
elseif p=2 then
HU="昨日链入排行"
gettxt="where active=0 and del=0 and DATEDIFF('d', HU_time, date())=1"
elseif p=3 then
HU="本周链入排行"
gettxt="where active=0 and del=0 and DATEDIFF('d', HU_time, date())<8"
elseif p=4 then
HU="月链入排行"
gettxt="where active=0 and del=0 and DATEDIFF('d', HU_time, date())<31"
else
HU="总链入排行"
gettxt="where active=0 and del=0"
end if
Response.write "<card title="""&HU&"""><p>"
Response.write"按链入|<a href='top.asp?act=out&amp;p="&p&"&amp;sid="&sid&"'>按链出</a><br/>"
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from 74hu_link "&gettxt&" order by HU_in desc"
else
if p=1 then
HU="今日链出排行"
gettxt="where active=0 and del=0 and DATEDIFF('d', outtime, date())=0"
elseif p=2 then
HU="昨日链出排行"
gettxt="where active=0 and del=0 and DATEDIFF('d', outtime, date())=1"
elseif p=3 then
HU="本周链出排行"
gettxt="where active=0 and del=0 and DATEDIFF('d', outtime, date())<8"
elseif p=4 then
HU="月链出排行"
gettxt="where active=0 and del=0 and DATEDIFF('d', outtime, date())<31"
else
HU="总链出排行"
gettxt="where active=0 and del=0"
end if
Response.write "<card title="""&HU&"""><p>"
Response.write"<a href='top.asp?act=in&amp;p="&p&"&amp;sid="&sid&"'>按链入</a>|按链出<br/>"
Set rs = Server.CreateObject("ADODB.Recordset")
sql="Select * from 74hu_link "&gettxt&"  order by [HU_out] desc"
end if
rs.open sql,conn,1,1
If Not rs.eof	Then
	Dim PageSize,i
	PageSize=15
	Dim Count,page,pagecount,gopage
	gopage="top.asp?act="&act&"&amp;p="&p&"&amp;sid="&sid&"&amp;"
	Count=rs.recordcount
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For
Response.write ""&i+(page-1)*PageSize&"."
Response.write "<a href=""Showlink.asp?class="&rs("classid")&"&amp;id="&rs("id")&"&amp;sid="&sid&""">"&usb(rs("name"))&"</a>(出"&rs("HU_out")&"/入"&rs("HU_in")&")<br/>"
	rs.moveNext
 	Next
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
if pagecount>1 then response.write "<br/>"&page&"/"&pagecount&"页,共"&Count&"条<br/><input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">跳到该页</a><br/>"

Else
Response.write "暂时没有！<br/>"
end if

Response.write "<br/><a href=""top.asp?act="&act&"&amp;p=1&amp;sid="&sid&""">今</a>|"
Response.write "<a href=""top.asp?act="&act&"&amp;p=2&amp;sid="&sid&""">昨</a>|"
Response.write "<a href=""top.asp?act="&act&"&amp;p=3&amp;sid="&sid&""">周</a>|"
Response.write "<a href=""top.asp?act="&act&"&amp;p=4&amp;sid="&sid&""">月</a>|"
Response.write "<a href=""top.asp?act="&act&"&amp;p=5&amp;sid="&sid&""">总</a><br/>"%>

<a href="Link_class.asp?sid=<%=sid%>">[分类管理]</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>[友链管理]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>
<%call ALLClose()%>