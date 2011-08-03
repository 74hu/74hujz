<!-- #include file="h.asp" -->
<%
IF  Request.QueryString("action")="view" Then
call view
else
call index
end if

sub index%>
<card title="最新公告"><p><%
Set Rs = Server.CreateObject("Adodb.Recordset")
Sql = "SELECT * FROM 74hu_gonggao order by id desc"
Rs.Open Sql,conn,1,1
If Not rs.eof Then
Dim PageSize,i
PageSize=10
Dim Count,page,pagecount,gopage	
gopage="?aid=gonggao&amp;"
Count=rs.recordcount
page=request.QueryString ("page")
if isnull(page) or isnumeric(page)=false then
page=1
end if
if page<1 then
page=1
end if
page=int(page)

if page<=0 or page="" then page=1
pagecount=(count+pagesize-1)\pagesize
if page>pagecount then page=pagecount	
rs.move(pagesize*(page-1))
response.write ("共:"&count&"条公告<br/>")
For i=1 To PageSize
If rs.eof Then Exit For	%>
<a href="?aid=gonggao&amp;action=view&amp;id=<%=rs("id")%>"><%=i+(page-1)*PageSize%>.<%=ubb(Rs("name"))%></a><br/> <%
rs.moveNext
Next
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">[GO]</a><br/>"
Else
response.write "暂时没有公告！<br/>"
end if
Rs.close
set Rs=nothing
end sub%>
<%sub view%>
<card id="index" title="查看公告"><p>
<%

Set rs=Server.CreateObject("Adodb.Recordset")
rs.open "select  * from 74hu_gonggao where id="&id&" order by id desc",conn,1,1
If Not rs.eof Then
response.write "标题:"&UBB(rs("name"))&"<br/>"
response.write "["&fordate(rs("HU_time"))&"]<br/>"
response.write "内容:"&ubbcode(rs("title"))&"<br/>"

Else
response.write "没有这个公告！"
end if
response.write "<br/><a href=""?aid=gonggao"">返回公告中心</a><br/>"

Rs.close
set Rs=nothing

end sub%>