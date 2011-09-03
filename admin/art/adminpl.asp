<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="评论管理">
<p>
 <%
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "Select * from 74hu_pl order by id desc",conn,1,1
If Not rs.eof	Then

	PageSize=10
	gopage="adminpl.asp?sid="&sid&"&amp;"
	Count=conn.execute("Select count(ID) from 74hu_pl")(0)
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
        if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize

smsid=rs("Smsid")
plid=rs("id")
plip=rs("ip")
plnr=rs("pl")
plsj=rs("pltime")

set rs1=server.createobject("adodb.recordset")
sql="Select classid,title from 74hu_article where ID="&smsid
rs1.open sql,conn,1,1

id = smsid
ids = rs1("classid")
title = rs1("title")

	If rs.eof Then Exit For
%><%=i+(page-1)*PageSize%>、评论文章：<a href="smsview.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;ids=<%=ids%>&amp;TP=1"><%=noubb(title)%></a><br/>评论者IP：<%=plip%><br/>时间：<%=fordate2(plsj)%><br/>内容：<%=noubb(plnr)%><br/><a href="delpl.asp?sid=<%=sid%>&amp;id=<%=plid%>">[删除这条评论]</a><br/>

<%
	rs.moveNext
	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a>"
Else
%>
	暂时没有评论！
<% end if%>
       <%
	rs.close
	set rs=nothing
	rs1.close
	set rs1=nothing
	conn.close
	set conn=nothing

%>
<br/>
<a href="qkpin.asp?sid=<%=sid%>">[清空评论]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>