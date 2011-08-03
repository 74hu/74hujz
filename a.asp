<!--#include file="h.asp"--><%
p=request.QueryString("p")
if IsNull(p) or p="" or IsNumeric(p)=False then p=1
Set rs = Server.CreateObject("ADODB.Recordset")
sql="Select title,test,hit,smspin,classid,HU_author,HU_date from 74hu_article where id="&id
rs.open sql,conn,1,3
if rs.eof then
rs.close
set rs=Nothing
response.redirect "?aid=index"
response.end
end if
ids=rs("classid")
Set rss = Server.CreateObject("ADODB.Recordset")
sql="Select class from 74hu_list where classid="&ids
rss.open sql,conn,3,1 
if rss.eof then
rss.close
set rss=Nothing
response.redirect "?aid=index"
response.end
end if
rs("hit")=rs("hit")+1
rs.update()
response.write"<card title='"&ubb(rs("title"))&"-"&rss("class")&"'><p>"
if wapgonggao="1" then
response.write "<a href='?aid=gonggao'><img src='images/msg.gif' alt='.'/>网站发布最新公告!</a><br/>" & Chr(13)
end if
response.write"-<a href='?aid=index'>首页</a>-<a href='?aid=list&amp;id="&ids&"&amp;page="&p&"'>"&rss("class")&"</a>-正文<br/>"
Counts=rs("smspin")
Response.Write""&ubb(rs("title"))&"<br/>-----------<br/>"
if adsetkf("ads1")=1 then
call adstr(1)
response.write"<br/>"
end if
Response.Write"内容来源:"&rs("HU_author")&"<br/>["&fordate(rs("HU_date"))&"]<br/>"
Content=rs("test")
pageWordNum=viewtnums
StartWord = 1
Length=len(Content)
PageAll=(Length+PageWordNum-1)\PageWordNum
page=request.QueryString("page")
if Isnull(page) or IsNumeric(page)=False or page="" then page=1
if page<1 then page=1
i=int(page-1)
if page>PageAll then page=PageAll
if isnull(i) or IsNumeric(i)=False then i=0
dim ccc,sss
ccc=instr(content,"||")
if ccc>0 then
sss=split(content,"||")
PageAll=ubound(sss)+1
if i>PageAll-1 then i=PageAll-1
content = sss(i)
else
if clng(i)>int(PageAll) then i=PageAll-1
Content = mid(Content,StartWord+i*PageWordNum,PageWordNum)
end if
response.write("" &ubbcode(content)& "") & vbnewline
if 0<=i<PageAll then
Response.Write "<br/>"
end if
if cint(i)<cint(PageAll)-1 then
Response.Write "<a href='?aid=art&amp;id=" & id & "&amp;page=" & i+2 & "&amp;p=" & p & "'>下页</a>"&"&nbsp;" 
End if
if cint(i)>0 then 
Response.Write "<a href='?aid=art&amp;id=" & id & "&amp;page=" & i & "&amp;p=" & p & "'>上页</a>"
End if
if PageAll>1 then
response.write "(" & i+1 & "/" & PageAll & ")"%><br/>第<input name="i" type="text" format="*N" emptyok="true" size="2" value="" maxlength="2"/>页
<anchor>跳转<go href="?aid=art&amp;id=<%=id%>&amp;p=<%=p%>" accept-charset='utf-8'><postfield name="page" value="$(i)"/></go></anchor><br/>
<%end if%>-----------<br/>※快速评论：<br/><input type="text" name="pl<%=minute(now)%><%=second(now)%>" title="输入内容" value="" maxlength="200"/><br/>
<anchor title="确定">提交<go method="get" href="?aid=diss&amp;id=<%=id%>&amp;p=<%=p%>">
<postfield name="pl" value="$(pl<%=minute(now)%><%=second(now)%>)"/><postfield name="ip" value="$(ip)"/></go></anchor> <a href='?aid=dis&amp;id=<%=id%>&amp;p=<%=p%>'>网友评论(<%=Counts%>)条</a><br/>
<%set rs1=server.createobject("adodb.recordset")
sql="select top 1 id,test,title from 74hu_article where classid="&ids&" and id<"&id&" order by id desc"
rs1.open sql,conn,3,1
if rs1.recordcount>0 then
%><a href="?aid=art&amp;id=<%=rs1("id")%>&amp;p=<%=p%>">&gt;&gt;<%=ubb(rs1("title"))%></a><br/>
<%end if
rs1.close
set rs1=nothing
set rs2=server.createobject("adodb.recordset")
sql="select top 1 id,test,title from 74hu_article where classid="&ids&" and id>"&id&" order by id asc"
rs2.open sql,conn,3,1
if rs2.recordcount>0 then
%><a href="?aid=art&amp;id=<%=rs2("id")%>&amp;p=<%=p%>">&lt;&lt;<%=ubb(rs2("title"))%></a><br/>
<%end if
response.write("[相关内容]<br/>")
call wendtitle(3,ids)
if adsetkf("ads2")=1 then
call adstr(2)
response.write"<br/>"&chr(13)
end if
response.write""&wapurl&" ["&fordate2(Now)&"]"
rs.close
set rs=nothing
rss.close
set rss=nothing
rs2.close
set rs2=nothing%>