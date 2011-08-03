<!--#include file="h.asp"-->
<%act=request.QueryString("act")
if act="add" then%>
<card title="申请友链"><p>
网站名称:(3-6字)<br/><input name="name<%=minute(now)%><%=second(now)%>" maxlength="7" value=""/><br/>
网站简称:(2汉字)<br/><input name="namt<%=minute(now)%><%=second(now)%>" maxlength="2" value=""/><br/>
网址:(需http://)<br/><input name="url<%=minute(now)%><%=second(now)%>" value="http://"/><br/>
网站分类：<select name="classid<%=minute(now)%><%=second(now)%>">
<%Set Rs=server.createobject("adodb.recordset")
Sql = "select classid,class from 74hu_linkc"
Rs.open Sql,conn,1,1
do while not Rs.eof
%><option value='<%=rs("classid")%>'><%=rs("class")%></option>
<%rs.movenext
Loop%></select><br/>
网站简介：(50字内)<br/><input name="jian<%=minute(now)%><%=second(now)%>" title="简介"  value="暂时没有介绍…" maxlength="100"/><br/><anchor>确定提交<go href="?aid=link&amp;act=post" method="get" accept-charset="utf-8">
<postfield name="name" value="$(name<%=minute(now)%><%=second(now)%>)"/>
<postfield name="namt" value="$(namt<%=minute(now)%><%=second(now)%>)"/>
<postfield name="url" value="$(url<%=minute(now)%><%=second(now)%>)"/>
<postfield name="classid" value="$(classid<%=minute(now)%><%=second(now)%>)"/>
<postfield name="jian" value="$(jian<%=minute(now)%><%=second(now)%>)"/>
</go></anchor><br/>
<br/>欢迎优秀WAP网站交换链接。
<br/>1.合作原则:流量互补,双赢发展,10天没流量首页自动隐藏。
<br/>2.流程: 
<br/>1)提交网站，获取链接地址; 
<br/>2)将我站的链接放到贵站明显位置。
<br/>3)我站人员3个工作日内审核网站，合适网站即可收录。
<br/>
<br/>申请友情链接前请先在您的网站上做好本站的链接：
<br/>网站名称：<%=waptitle%>
<br/>做好我站链接后，我们会及时进行审核。<br/>
<%Response.write "<a href='?aid=link'>返回友链首页</a><br/>"
rs.close
set rs=nothing
elseif act="go" then
On Error Resume Next
Server.ScriptTimeOut=9999999
yourip=Request.ServerVariables("HTTP_X_UP_CALLING_LINE_ID")
if yourip="" then yourip=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
if yourip="" then yourip=Request.ServerVariables("REMOTE_ADDR")
sss=180
ips=500

cache_ip=Application("cache_ip")
if cache_ip="" then cache_ip="|"
one_ip=split(cache_ip,"|")
all_s=ubound(one_ip)
for k_ip=0 to all_s
if yourip=one_ip(k_ip) then
i_ip=k_ip:ip_time=one_ip(k_ip+1):Exit for
else
i_ip=0:ip_time="2000-10-10 10:10:10"
end if
next
del_time=DATEDIFF("s",ip_time,now())

if i_ip<all_s and sss>del_time then
response.redirect "?aid=index"
response.end
end if
if all_s>ips*2 Then
Application.Lock
Application("cache_ip")="|"
Application.UnLock
else
if i_ip=0 then
temp_s=cache_ip&yourip&"|"&now()&"|"
else
text_1="|"&yourip&"|"&ip_time&"|"
num_1=len(cache_ip)
num_2=len(text_1)
num_3=instr(cache_ip,text_1)
num_4=num_1-num_2-num_3+1
text_2=left(cache_ip,num_3)
text_3=right(cache_ip,num_4)
text_4=yourip&"|"&now()&"|"
temp_s=text_2&text_4&text_3
end if
Application.Lock
Application("cache_ip")=temp_s
Application.UnLock
end if
set Rs=Server.CreateObject("ADODB.Recordset")
Sql="select ID,HU_in,HU_time from 74hu_link Where ID="&ID
Rs.open Sql,conn,1,3
Rs("HU_in")=Rs("HU_in")+1
Rs("HU_time")=now()
Rs.update()
rs.close
set rs=nothing
response.redirect "?aid=index"
response.end
elseif act="view" then
set Rs=Server.CreateObject("ADODB.Recordset")
Sql="select * from 74hu_link Where id="&id
Rs.open Sql,conn,1,3
If Not rs.eof	Then
Rs("HU_out")=Rs("HU_out")+1
Rs("OUTtime")=now()
Rs.update()
Else
Response.Write	("该网站不存在")
End If%>
<card title='<%=usb(rs("name"))%>' ontimer='<%=ubb(rs("url"))%>'><timer value='1'/><p>
正在跳转到“<%=usb(rs("name"))%>”,<br/>请稍候...<a href='<%=ubb(rs("url"))%>'>快速进入</a><br/>
网站介绍：<%=usb(rs("jian"))%><br/>

<br/>
<%rs.close
set rs=nothing
elseif act="post" then
classid=Request.QueryString("classid")
name=hu(Request.QueryString("name"))
namt=hu(Request.QueryString("namt"))
url=LCase(hu(Request.QueryString("url")))
jian=hu(Request.QueryString("jian"))
if session("name")=1 then
nolink
else
if name="" or namt="" or url="" or jian="" or classid="" or isnumeric(classid)=false then
noname
else
Set RSS=server.createobject("adodb.recordset")
Sqll="select * from 74hu_ad"
RSS.open sqll,conn,1,1
active=rss("active")
rss.close
set rss=nothing
Set RS=server.createobject("adodb.recordset")
Sql="select * from 74hu_link"
RS.open sql,conn,1,3
RS.addnew
RS("name")=name
RS("namt")=namt
RS("url")=url
RS("classid")=classid
RS("jian")=jian
RS("active")=active
RS.update
session.timeout=1
session("name")=1
end if
end if
response.write "<card title='申请友链成功' ontimer='?aid=link&amp;act=you'><timer value='1'/><p>"%>
申请友链成功<br/>
<%sub nolink()%>
<card title="重复申请"><p>
你刚才已申请过了！请不要重复申请！<br/>
<%Response.End
End Sub
sub noname()%>
<card title="出错了吧"><p>
各项都要填写,不能为空！<br/>
<%Response.End
End Sub
rs.close
set rs=nothing
elseif act="list" then
add=request.QueryString("class")
if add="" or IsNumeric(add)=false then
response.redirect"?aid=index"
response.end
else
Set rss=Server.CreateObject("ADODB.Recordset")
sqll="Select * from 74hu_linkc where classid="&add
rss.open sqll,conn,1,1 
if not rss.eof then
classname=rss("class")
end if
rss.close
set rss=nothing%>
<card title="<%=classname%>网站"><p> 
<%Response.write "=" &classname&"网站=<br/>"
Set rs = Server.CreateObject("ADODB.Recordset")
sql="Select * from 74hu_link where classid="&add&" And Active=0 and del=0  order by HU_time desc"
rs.open sql,conn,1,1 
If Not rs.eof	Then
PageSize=15
gopage="?aid=link&amp;act=list&amp;class="&ADD&"&amp;"
Count=conn.execute("Select count(ID) from 74hu_link where classid="&add&"")(0)
page=request("page")
if page="" or isnumeric(page)=false then page=1
page=int(page)
if page<=0 then page=1
pagecount=(count+pagesize-1)\pagesize
if page>pagecount then page=pagecount
rs.move(pagesize*(page-1))
For i=1 To PageSize
If rs.eof Then Exit For
Response.write ""&i+(page-1)*PageSize&"."
Response.write "<a href=""?aid=link&amp;act=view&amp;class="&rs("classid")&"&amp;id="&rs("id")&""">"&usb(rs("name"))&"</a><br/>"
rs.moveNext
Next
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a><br/>"
Else
%>暂时没有添加！<br/>
<%end if%><br/><a href='?aid=link'>返回友链首页</a><br/>
<%rs.close
set rs=nothing
end if
elseif act="you" then%>
<card title="申请友链成功"><p>
<%set Rss=Server.CreateObject("ADODB.Recordset")
Sqll="select Active from 74hu_ad"
Rss.open Sqll,conn,1,1
Active=Rss("Active")
rss.close
set rss=nothing
set Rs=Server.CreateObject("ADODB.Recordset")
Sql="select top 1 * from 74hu_link order by id desc"
RS.open Sql,conn,1,1%>
添加友链地址成功，<% if Active=1 then%>请等待站长审核，审核通过后才会显示<br/>
<%else%>你的友链已经显示出来!<br/>
<%End if%>
贵站返回我站的链接地址是:http://<%=wapurl%>/?aid=link&amp;act=go&amp;id=<%=Rs("ID")%><br/>
网站名称:<%=waptitle%><br/>
<%rs.close
set rs=nothing
else%><card title="友情链接"><p>----动态友链----<br/>
<%Set rs = Server.CreateObject("ADODB.Recordset")
sql="select linkactive from 74hu_ad"
rs.open sql,conn,1,1
if not rs.eof then
getday=rs("linkactive")
end if
rs.close
set rs=nothing

set rs=Server.CreateObject("ADODB.Recordset")
Sql="select ID,name,classid from 74hu_link Where active=0 and del=0 and datediff('d', HU_time, now())<"&getday&" order by HU_time desc"
rs.open Sql,conn,1,1
If Not rs.eof	Then

PageSize=30
gopage="?aid=link&amp;"
Count=rs.recordcount
page=request("page")
if page="" or isnumeric(page)=false then page=1
page=int(page)
if page<=0 then page=1
pagecount=(count+pagesize-1)\pagesize
if page>pagecount then page=pagecount
rs.move(pagesize*(page-1))

For i=1 To PageSize
If rs.eof Then Exit For
Response.write "<a href=""?aid=link&amp;act=view&amp;id="&rs("id")&"&amp;class="&rs("classid")&""">"&usb(rs("name"))&"</a><br/>" 
rs.moveNext
Next
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">跳转</a>"
Else%>
暂时没有友链！<br/>
<%end if%>
<%rs.close
set rs=nothing%>
----网站分类----<br/>
<%dim linkindex
set rs = server.createobject("adodb.recordset")
sql="select linkindex from 74hu_ad"
rs.open sql,conn,1,1
if rs.eof then
rs.close
set rs=nothing
response.write("<p>资料没有配置！</p></card></wml>")
response.end
end if
linkindex=rs("linkindex")
rs.close
set rs=nothing
if linkindex=0 then
set rs1 = server.createobject("adodb.recordset")
sql1="select * from 74hu_linkc order by pid asc"
rs1.open sql1,conn,1,1
if rs1.eof then
response.write("暂无分类<br/>")
else
i=1
do while not rs1.eof
set Rss=Server.CreateObject("ADODB.Recordset")
Sqls="select top 4 ID,namt,classid from 74hu_link Where Active=0 and del=0 and classid="&rs1("classid")&" order by HU_time desc"
rss.open Sqls,conn,1,1
Response.write "【<a href=""?aid=link&amp;act=list&amp;class="&rs1("classid")&""">"&usb(rs1("Class"))&"</a>】"
If rss.eof then
Response.write "暂时还没有<br/>"& chr(13)
Else
for i=1 to 6
Response.write "<a href='?aid=link&amp;act=view&amp;class="&rss("classid")&"&amp;id="&rss("id")&"'>"&usb(rss("namt"))&"</a>"& chr(13) 
rss.Movenext
if rss.EOF then Exit for
Next
Response.write "<br/>"& chr(13) 
End if
rss.close
set rss=nothing
i=i+1
rs1.movenext
loop
end if
rs1.close
set rs1=nothing
else
set rs = server.createobject("adodb.recordset")
sql="select * from 74hu_linkc order by pid asc"
rs.open sql,conn,1,1
if not (rs.bof and rs.eof)  then
For i=1 to rs.RecordCount
If Rs.Eof Then
exit For
End If
if rs("br")="1" then
br="<br/>"
else
br=""
end if
Response.write "<a href=""?aid=link&amp;act=list&amp;class="&rs("classid")&""">"&usb(rs("class"))&"</a> "& br &"" 
Rs.MoveNext
Next
end if
rs.close
set rs=nothing
end if
Response.write "<a href='?aid=link&amp;act=add'>&gt;&gt;友链合作申请</a>"& chr(13)
end if%>