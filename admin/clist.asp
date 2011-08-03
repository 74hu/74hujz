<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="栏目分类管理">
<p><%=now%><br/>
<% id= int(request.QueryString("id"))
   idd= int(request.QueryString("idd"))
   lxl= request.QueryString("lx")
if lxl=0 then %>
<a href="classadd.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;lxl=<%=lxl%>">[添加栏目]</a><br/>
<%end if%>
<%
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_class where classid="&id,conn,1,1

if rs("lx")=0 then lxx="页面菜单"
if rs("lx")=1 then lxx="文章栏目"
if rs("lx")=8 then lxx="随机广告"
if rs("lx")=9 then lxx="WML标签"
if rs("lx")=10 then lxx="最新文章"
if rs("lx")=11 then lxx="最热文章"
if rs("lx")=12 then lxx="随机文章"
if rs("lx")=19 then lxx="站内搜框"
if rs("lx")=20 then lxx="WML页面"
response.write ("项目名称:"&ubb(left(rs("class"),30))&"<br/>")
response.write ("项目类型:"&lxx&"<br/>")
response.write ("项目编号:"&rs("lx")&"<br/>")
response.write ("项目排序:"&rs("pid")&"<br/>")
if rs("br")=1 then
response.write ("项目换行:换行<br/>")
else
response.write ("项目排序:不换行<br/>")
end if
response.write ("-----------<br/>")
rs.close
set rs=nothing
%>
<% set rs=server.createobject("adodb.recordset")
sql = "select * from 74hu_class where parent="&id&" order by pid asc"
rs.open sql,conn,1,1
If Not rs.eof	Then
	Dim PageSize,i
	PageSize=20
	gopage="clist.asp?sid="&sid&"&amp;id="&id&"&amp;"
	Count=rs.recordcount
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
        if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For
%><a href="clist.asp?sid=<%=sid%>&amp;id=<%=rs("classid")%>&amp;idd=<%=id%>&amp;lx=<%=rs("lx")%>&amp;lxl=<%=lxl%>">[管理]></a><%=rs("pid")%>.<%=ubb(left(rs("class"),20))%><br/>

<%     
	rs.moveNext
	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a>"
Else
%>
	(没有下级栏目)

<%end if%>

<br/><a href="editclass.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;lx=<%=lxl%>">[编辑栏目]</a><br/>
<a href='move.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;lx=<%=lxl%>'>[移动栏目]</a><br/>
<a href='delcl.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;idd=<%=idd%>&amp;lx=<%=lxl%>'>[删除栏目]</a><br/>
<%if idd<>0 then %>
<a href="Clist.asp?sid=<%=sid%>&amp;id=<%=idd%>&amp;lx=0">[栏目分类]</a><br/>
<%end if%>
<a href="class.asp?sid=<%=sid%>">[设计中心]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%><br/>
</p></card></wml>