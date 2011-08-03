
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<%dim rs1,sql1,id
id= int(request.QueryString("id"))
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "Select * from 74hu_list where classid="&id,conn,1,1 
if not rs1.eof then
classname=rs1("class")
end if
rs1.close
set rs1=nothing
%>
<card title="<%=classname%>"><p>
<%
dim rs,sql
response.write ("<a href='add.asp?sid="&sid&"&amp;id="&id&"'>[添加文章]</a><br/>")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="Select * from 74hu_article where classid="&id&" order by id desc"
rs.open sql,conn,1,1
 If Not rs.eof	Then

	PageSize=10
	gopage="adminsmscl.asp?sid="&sid&"&amp;ID="&ID&"&amp;"
	Count=conn.execute("Select count(ID) from 74hu_article where classid="&id&"")(0)
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
        if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize

	If rs.eof Then Exit For
%><a href='wzgl.asp?sid=<%=sid%>&amp;id=<%=rs("id")%>&amp;classid=<%=rs("classid")%>'>[管理]</a><%=i+(page-1)*PageSize%>.<a href='smsview.asp?sid=<%=sid%>&amp;id=<%=rs("id")%>&amp;ids=<%=rs("classid")%>'><%=ubb(rs("title"))%></a><br/>
<%     
	rs.moveNext
	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a>"
Else
%>
	暂时没有文章！
       <%
        end if
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing

%>
<br/>----------<br/>
<a href='upfile.asp?sid=<%=sid%>&amp;id=<%=id%>'>[2.0传文章]</a><br/>
<a href="wzclass.asp?sid=<%=sid%>">[文章分类]</a><br/>

<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>