<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title='文章分类'><p>
请选择要移动的分类:<br/>
<%id=int(request.QueryString("id"))
CLASSid=int(request.QueryString("CLASSid"))
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "SELECT * from 74hu_list order by classid desc",conn,1,1
	If Not rs.eof	Then
	PageSize=10
	gopage="move.asp?sid="&sid&"&amp;id="&id&"&amp;"
	Count=rs.recordcount
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
        if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For
%><a href="wzmove.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;classid=<%=rs("classid")%>"><%=ubb(rs("class"))%></a><br/>
<%     
	rs.moveNext
	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a>"
Else
%>
	(没有栏目)

<%end if%><br/>
<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=classid%>">[文章列表]</a><br/>
<a href="wzclass.asp?sid=<%=sid%>">[文章分类]</a><br/>

<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>