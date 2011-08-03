<!-- #include file="h.asp" -->
<%p=request.QueryString("p")
if IsNull(p) or IsNumeric(p)=False or p="" then p=1%><card title='网友跟贴'><p>-<a href='?aid=index'>首页</a>-<a href='?aid=art&amp;id=<%=id%>&amp;p=<%=p%>'>查看原文</a>-跟贴
<br/>发表评论：<br/>
<input type="text" name="pl<%=minute(now)%><%=second(now)%>" title="输入内容" value="" maxlength="200"/><br/>
<anchor title="确定">提交
<go method="get" href="?aid=diss&amp;id=<%=id%>&amp;p=<%=p%>">
<postfield name="pl" value="$(pl<%=minute(now)%><%=second(now)%>)"/><postfield name="ip" value="$(ip)"/>
</go></anchor><br/>【网友评论区】<br/><%
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select id,pl,pltime,ag,da from 74hu_pl where smsid="&id&" order by id desc",conn,3,1
If Not rs.eof	Then
PageSize=8
gopage="?aid=dis&amp;id="&id&"&amp;p="&p&"&amp;"
Count=conn.execute("Select count(ID) from 74hu_pl where smsid="&id&"")(0)	
page=request.QueryString("page")
if page="" or isnumeric(page)=false then page=1
page=int(page)
if page<=0 then page=1
pagecount=(count+pagesize-1)\pagesize
rs.move(pagesize*(page-1))
For i=1 To PageSize
If rs.eof Then Exit For	
response.write ""&fordate2(rs("pltime"))&"发表<br/>　　"   
response.write rs("pl")&"<br/>-----------<br/>"
rs.moveNext
Next
if page>1 then response.write "<a href="""&gopage&"page=1"">首页</a>&nbsp;"
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>&nbsp;"
if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>&nbsp;"
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&pagecount&""">末页</a>"
if pagecount>1 then response.write "<br/>第"&page&"页 共"&pagecount&"页<br/>第<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/>页 <a href="""&gopage&"page=$(page)"">跳转</a><br/>"
Else
response.write"暂时没有评论！<br/> "
end if
rs.close
set rs=nothing%><a href='?aid=art&amp;id=<%=id%>&amp;p=<%=p%>'>返回原文</a> <a href='?aid=list&amp;p=<%=p%>'>返回上级栏目</a><br/>