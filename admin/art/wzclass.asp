
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="文章分类管理">
<p>
<%
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_list order by classid asc",conn,1,1
If Not rs.eof	Then
	PageSize=15
	gopage="wzclass.asp?sid="&sid&"&amp;"
	Count=rs.recordcount
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
        if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For
%>[<a href="wzclasscl.asp?sid=<%=sid%>&amp;id=<%=rs("classid")%>">管理</a>]<%=i+(page-1)*PageSize%>.<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=rs("classid")%>"><%=ubb(left(rs("class"),10))%></a>[栏目ID<%=rs("classID")%>]<br/>
<%     
	rs.moveNext
	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a>"
Else
%>
	暂时没有文章栏目！

<%
end if

rs.close
set rs=nothing
%><br/>-----------<br/>
<a href="tjwz.asp?sid=<%=sid%>">[添加分类]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>