<!--#include file="Head.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="mymin.asp"-->
<card title="全部友链"><p>
<%
dim Rs,Sql,idd
idd=Request.Querystring("id")
call conndata
set Rs=Server.CreateObject("ADODB.Recordset")
Sql="select * from 74hu_link order by HU_time desc"
rs.open Sql,conn,1,1
If Not rs.eof	Then
Dim PageSize,i
	PageSize=10
	Dim Count,page,pagecount,gopage	
	gopage="links.asp?sid="&sid&"&amp;"
	Count=conn.execute("Select count(ID) from 74hu_link")(0)
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
        if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
        response.write ("共:"&count&"条友链<br/>")
	For i=1 To PageSize
	If rs.eof Then Exit For	
        Response.write "<a href=""Showlink.asp?sid="&sid&"&amp;class="&rs("classid")&"&amp;id="&rs("id")&""">"&usb(rs("name"))&"</a>(出"&rs("HU_out")&"/入"&rs("HU_in")&")<br/>"
Response.write "<a href='"&UBB(rs("URL"))&"'>查看</a>|"
Response.write "<a href='Edit_link.asp?sid="&sid&"&amp;class="&rs("classid")&"&amp;id="&rs("id")&"'>编辑</a>|"
if rs("del")="0" then
Response.write "<a href='DEL_link.asp?sid="&sid&"&amp;class="&rs("classid")&"&amp;id="&rs("id")&"&amp;TP=2'>删除</a><br/>"
else
Response.write "<a href='hui_link.asp?sid="&sid&"&amp;class="&rs("classid")&"&amp;id="&rs("id")&"&amp;TP=2'>恢复</a><br/>"
end if
	rs.moveNext
	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">跳转</a><br/>"
Else
Response.write ("该分类暂时没有友链！")
end if%><br/>
<a href="Link_class.asp?sid=<%=sid%>">[分类管理]</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>[友链后台]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>
<%call ALLClose()%>