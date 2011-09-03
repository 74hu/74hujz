<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="网站设计中心">
<p>
<%
IF KEY<>0 then
  Call Error("你的权限不足！")
  end if
%>
<a href="classadd.asp?sid=<%=sid%>">[添加栏目]</a>
<br/>-----------<br/>
<%
	rs.close
	set rs=nothing
	conn.Close
	set conn=nothing
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_class where parent=0 order by pid asc",conn,1,1

If Not rs.eof	Then
	PageSize=20
	gopage="class.asp?sid="&sid&"&amp;"
	Count=rs.recordcount
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
        if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For
%><a href="clist.asp?sid=<%=sid%>&amp;id=<%=rs("classid")%>&amp;lx=<%=rs("lx")%>">[管理]></a><%=rs("pid")%>.<%=noubb(left(rs("class"),10))%><br/>
<%     
	rs.moveNext
	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a><br/>"
Else
%>
	暂时没有栏目！<br/>

<%
end if

rs.close
set rs=nothing
call CloseConn
%>-----------<br/>
<a href="classadd.asp?sid=<%=sid%>">[添加栏目]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
</p></card></wml>
