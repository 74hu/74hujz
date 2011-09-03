<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<%
p=cint(request.QueryString("p"))
if p="" or p<1 then p=1
%>
<card title="留言管理">
<p align="left">
<%
call conndata
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from 74hu_guest order by ID desc",conn,1,1
If Not rs.eof	Then
	PageSize=10
	j=0
	gopage="index.asp?sid="&sid&"&amp;p="&p&"&amp;"
	Count=conn.execute("Select count(ID) from 74hu_guest")(0)
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount	
	rs.move(pagesize*(page-1))
	Response.write ""&now()&"<br/>"
	response.write ("共"&count&"条留言<br/>")
	For i=1 To PageSize
	If rs.eof Then Exit For
	j=j+1%>
[<a href="delly.asp?sid=<%=sid%>&amp;ID=<%=rs("ID")%>&amp;p=<%=p%>">删除</a>]<%=j+(page-1)*PageSize%>.<a href='readd.asp?sid=<%=sid%>&amp;ID=<%=rs("ID")%>&amp;p=<%=p%>'><%=noubb(rs("title"))%></a><% if rs("retext")<>"" then %>[已回]<% end if%><br/>
<%
	rs.moveNext
	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">[GO]</a><br/>"
Else
Response.write ("暂时没有留言！<br/>")
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
----------<br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a><br/>
</p></card></wml>