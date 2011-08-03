<!--#include file="Head.asp"-->

<!--#include file="Config.asp"-->
<!--#include file="mymin.asp"-->
<card title="管理友链类别"><p>
<% dim rs,Sql
call conndata
set rs=server.createobject("adodb.recordset")
sql = "select * from 74hu_linkc Order by pid asc"
rs.open sql,conn,1,1
If Not rs.eof	Then
	Dim PageSize,i
	PageSize=20					
	Dim Count,page,pagecount,gopage			
	gopage="Link_class.asp?sid="&sid&"&amp;"
	Count=rs.recordcount
        response.write (""&now()&"<br/>")
        response.write ("共:"&count&"条友链分类<br/>")
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1		
	pagecount=(count+pagesize-1)\pagesize	
        if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))					
	For i=1 To PageSize     						
	If rs.eof Then Exit For	%>			
<a href="mymin_class.asp?sid=<%=sid%>&amp;id=<%=rs("classid")%>"><%=rs("class")%></a>|<a href="editclass.asp?sid=<%=sid%>&amp;id=<%=rs("classid")%>">修改</a>|<a href='delcl.asp?sid=<%=sid%>&amp;id=<%=rs("classid")%>'>删除</a><br/>   
	<%rs.moveNext
 	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a><br/>"
Else
%>
	暂时没有友链类别！<br/>
<%end if%>----------<br/>
<a href='mymin_index.asp?sid=<%=sid%>'>[友链管理]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml><%call ALLClose()%>

