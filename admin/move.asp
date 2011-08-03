<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card id="index" title="栏目移动管理"><p>
<%IF KEY<>0 then
  Call Error("你的权限不足！")
  end if%>
请选择要移动到的栏目:<br/><br/>
<%id= request.QueryString("id")
if id="" or IsNumeric(id)=False then
  Call Error("ID无效！")
  end if
  ids= request.QueryString("ids")
if ids="" or IsNumeric(ids)=False then ids=0
   lxl= request.QueryString("lx")

call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_class Where parent="&ids&" and lx=0",conn,1,1

If Not rs.eof Then
	PageSize=10
	gopage="move.asp?sid="&sid&"&amp;id="&id&"&amp;ids="&ids&"&amp;"
	Count=rs.recordcount
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
        if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For
%>
<a href="move.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;ids=<%=rs("classid")%>&amp;lx=<%=lxl%>">[+]</a>|<a href="movecl.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;lx=<%=lxl%>&amp;iid=<%=rs("classid")%>"><%=ubb(left(rs("class"),8))%></a><br/>
<%     
	rs.moveNext
	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a>"
Else
%>
	(没有栏目)

<%end if
rs.close
set rs=nothing
call CloseConn%><br/>
<a href="movecl.asp?sid=<%=sid%>&amp;iid=0&amp;id=<%=id%>&amp;lx=<%=lxl%>">[移到首页]</a><br/>
----------<br/>
<a href="class.asp?sid=<%=sid%>">[栏目管理]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>