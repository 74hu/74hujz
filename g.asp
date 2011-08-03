<!-- #include file="h.asp" -->
<%p=request.QueryString("p")
if p="" or isnumeric(p)=false then p=1
if p<1 then p=1
act=request.QueryString("act")%>
<card title="客服留言">
<%if act="view" then
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from 74hu_guest where ID=" & ID,conn,1,1
set rsn=Server.CreateObject("ADODB.Recordset")
rsn.open"select * from 74hu_guest where ID<"& ID &" order by id desc",conn,1,1
set rspr=Server.CreateObject("ADODB.Recordset")
rspr.open"select * from 74hu_guest where ID>"& ID &" order by id asc",conn,1,1
response.write("<p>-<a href='?aid=index'>首页</a>-<a href='?aid=guest'>客服</a>-查看留言<br/><br/>")
if rs.EOF then
response.write("无此留言！<br/>")
else
response.write"作者："&ubb(rs("name"))&"<br/>"
response.write(ubb(rs("text")) & "<br/>")
response.write("时间：" & fordate(rs("HU_time")) & "<br/>")

if rs("retext")<>"" then
response.write"----------<br/>"
response.write"回复："&ubb(rs("retext"))&"<br/>"
response.write"时间："&fordate(rs("retime"))&"<br/>"
end if
if rsn.recordcount>0 then response.write("<a href='?aid=guest&amp;act=view&amp;id=" & rsn("ID") & "&amp;p=" & p & "'>下条</a>&nbsp;")
if rspr.recordcount>0 then response.write("<a href='?aid=guest&amp;act=view&amp;id=" & rspr("ID") & "&amp;p=" & p & "'>上条</a>")
if rsn.recordcount>0 or rspr.recordcount>0 then response.write("<br/>")
rsn.close
set rsn=nothing
rspr.close
set rspr=nothing
end if

elseif act="add" then
randomize timer
ss=Int((9999)*Rnd +1000)
response.write"<p>-<a href='?aid=index'>首页</a>-<a href='?aid=guest'>客服</a>-发表留言<br/><br/>昵称：<br/>"
response.write"<input name=""name"" type=""text"" format=""*M"" emptyok=""true"" maxlength=""10""/><br/>"
response.write"标题：<br/><input name=""title"" type=""text"" format=""*M"" emptyok=""true"" maxlength=""20""/><br/>"
response.write"内容：<br/><input name=""text"" type=""text"" format=""*M"" emptyok=""true"" maxlength=""500""/><br/>"
response.write"联系方式(不公开)：<br/><input name=""lianxi"" type=""text"" format=""*M"" emptyok=""true"" maxlength=""50""/><br/>"
response.write"验证码："&ss&"<br/><input name=""num"" type=""text"" format=""*M"" emptyok=""true"" maxlength=""50""/><br/>"
response.write"<anchor>[提交留言]<go href=""?aid=guest&amp;act=save"" method=""get"" accept-charset=""utf-8"">"
response.write"<postfield name=""name"" value=""$(name)""/>"
response.write"<postfield name=""title"" value=""$(title)""/>"
response.write"<postfield name=""text"" value=""$(text)""/>"
response.write"<postfield name=""lianxi"" value=""$(lianxi)""/>"
response.write"<postfield name=""open"" value=""$(open)""/>"
response.write"<postfield name=""num"" value=""$(num)""/>"
response.write"<postfield name=""num1"" value="""&ss&"""/>"
response.write"</go></anchor><br/>"

elseif act="save" then

num=request.QueryString("num")
num1=request.QueryString("num1")

if num<>num1 then
response.write("<p>验证码错误,请返回重试！</p></card></wml>")
response.end
end if

name=hu(request.QueryString("name"))
title=hu(request.QueryString("title"))
text=hu(request.QueryString("text"))
lianxi=hu(request.QueryString("lianxi"))
name=ubb(name)
title=ubb(title)
text=ubb(text)
lianxi=ubb(lianxi)

if name="" or title="" or text="" then
response.write("<p>昵称或标题内容不能为空！</p></card></wml>")
response.end
end if
response.write"<onevent type='onenterforward'><go href='?aid=guest'/></onevent><p>"

set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from 74hu_guest",conn,1,2
rs.addnew
rs("name")=name
rs("title")=title
rs("text")=text
rs("HU_time")=now()
if lianxi<>"" then rs("lianxi")=lianxi
rs("agent")=User_Ip
rs.update
rs.close
set rs=Nothing

response.write"发表成功，正在返回！<br/>"
else
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from 74hu_guest order by id desc",conn,1,1
If Not rs.eof	Then
PageSize=10
gopage="?aid=guest&amp;"
Count=rs.recordcount
page=request.QueryString("page")
if page="" or isnumeric(page)=false or isnull(page) then page=1
page=int(page)
if page<=0 or page="" then page=1
pagecount=(count+pagesize-1)\pagesize
if page>pagecount then page=pagecount
rs.move(pagesize*(page-1))
response.write ("<p>-<a href='?aid=index'>首页</a>-客服首页<br/><br/>共"&count&"条")
response.write ("<a href=""?aid=guest&amp;act=add"">留言</a><br/>")
For i=1 To PageSize
If rs.eof Then Exit For
response.write "<a href='?aid=guest&amp;act=view&amp;id="&rs("ID")&"&amp;p="&p&"'>"
response.write ""&i+(page-1)*PageSize&"."&ubb(rs("title"))&"</a><br/>"
response.write "[网友:"&rs("name")
if rs("retext")<>"" then 
response.write "/已回"
else
response.write "/未回"
end if
response.write "]<br/>"
rs.moveNext
Next
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>&nbsp;"
if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
if pagecount>1 then response.write "<br/>"&page&"/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">跳转</a>"
Else
response.write "<p>还没有留言！<br/>"
end if
rs.close
set rs=nothing
response.write "<br/><a href=""?aid=guest&amp;act=add"">我要发表留言</a><br/>"
end if%>