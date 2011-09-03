<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<%
dim tit
tit=tt&minute(time)
tit=tt&second(time)

IF KEY<>0 then
  Call Error("<card title=""出错""><p>你的权限不足！")
  end if
call conndata
		IF  Request.QueryString("Action")="view" Then
			call view
		elseiF  Request.QueryString("Action")="del" Then
			call del
		elseiF  Request.QueryString("Action")="edit" Then
			call edit
		elseiF  Request.QueryString("Action")="add" Then
			call add
		elseiF  Request.QueryString("Action")="save" Then
			call save
		else

			call index
		end if

Function index%>
		<card id="index" title="公告管理">
		<p align="left">
		<%

Set Rs = Server.CreateObject("Adodb.Recordset")
Sql = "SELECT * FROM 74hu_gonggao order by id desc"
Rs.Open Sql,conn,1,1
If Not rs.eof	Then
	Dim PageSize,i
	PageSize=10							
	Dim Count,page,pagecount,gopage,parent			
	gopage="gonggo.asp?sid="&sid&"&amp;"
	Count=rs.recordcount
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1		
	pagecount=(count+pagesize-1)\pagesize	
        if page>pagecount then page=pagecount	
	rs.move(pagesize*(page-1))				
	For i=1 To PageSize       						
	If rs.eof Then Exit For				
%><%=i+(page-1)*PageSize%>.<a href='gonggo.asp?sid=<%=sid%>&amp;Action=view&amp;id=<%=noubb(rs("id"))%>'><%=noubb(rs("name"))%></a><br/>
<%     
	rs.moveNext						
	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a>"
Else
%>
	(没有公告)

<%end if%>
		<%
		Rs.close
		set rs=nothing
		Response.Write("<br/><a href='gonggo.asp?sid="&sid&"&amp;Action=add'>[发布公告]</a><br/>")
end Function
Function add%>
		<card id="add" title="发布公告"><p>
		标题:<br/><input name="name<%=tit%>" maxlength="255" /><br/>
		内容:(支持<a href="ubbcl.asp?sid=<%=sid%>">UBB</a>)<br/><input name="title<%=tit%>" /><br/> 
		<anchor>发布
		<go href="gonggo.asp?sid=<%=sid%>&amp;Action=save&amp;edit=add" method="post">
		<postfield name="name" value="$(name<%=tit%>)" />
		<postfield name="title" value="$(title<%=tit%>)" />
		</go>
		</anchor><br/>
<a href="gonggo.asp?sid=<%=sid%>">[公告管理]</a><br/>
	<%
end Function
Function edit
	id=TRim(Request("id"))
	if not isnumeric(id) then id=""
	if id=""  then%>
	<card id="index" title="出错啦"><p>ID无效.<br/>
	<%else

			Set Rs = Server.CreateObject("Adodb.Recordset")
			Sql = "SELECT * FROM 74hu_gonggao where id="&id
			Rs.Open Sql,conn,1,1
if not (rs.bof and rs.eof)  then %>
<card id="add" title="发布公告"><p>
标题:<br/><input name="name<%=tit%>" value="<%=rs("name")%>" maxlength="255" /><br/>
内容:(支持<a href="ubbcl.asp?sid=<%=sid%>">UBB</a>)<br/><input name="title<%=tit%>" value="<%=rs("title")%>" /><br/> 
	<anchor>修改
<go href="gonggo.asp?sid=<%=sid%>&amp;Action=save&amp;edit=edit" method="post">
			<postfield name="name" value="$(name<%=tit%>)" />
			<postfield name="id" value="<%=id%>" />
			<postfield name="title" value="$(title<%=tit%>)" />
			</go>
			</anchor><br/>
<a href="gonggo.asp?sid=<%=sid%>">[公告管理]</a><br/>
			<%else%>
			<card id="index" title="出错啦"><p>没有该公告!<br/>
			<%end if
			Rs.close
			set rs=nothing
		end if
end Function
Function view
	id=TRim(Request("id"))
	if not isnumeric(id) then id=""
		if id=""  then%>
	<card id="index" title="出错啦"><p>ID无效.<br/>
		<%else%>
		<card id="index" title="公告管理">
		<p>
		<%

			Set Rs = Server.CreateObject("Adodb.Recordset")
			Sql = "SELECT * FROM 74hu_gonggao where id="&id
			Rs.Open Sql,conn,1,1
			if not (rs.bof and rs.eof)  then %>
				标题:<%=noubb(rs("name"))%><br/>
				内容:<%=noubb(rs("title"))%><br/>
				时间:<%=noubb(rs("HU_time"))%><br/>
				<a href='gonggo.asp?sid=<%=sid%>&amp;Action=edit&amp;id=<%=id%>'>[编辑]</a> <a href='gonggo.asp?sid=<%=sid%>&amp;Action=del&amp;id=<%=id%>'>[删除]</a><br/>
<a href="gonggo.asp?sid=<%=sid%>">[公告管理]</a><br/>
			<%else%>
				暂无公告!<br/>
			<%end if
			Rs.close
			set rs=nothing
		end if
end Function
Function  del
	id=TRim(Request("id"))
	if not isnumeric(id) then id=""
 		if id=""  then%>
	<card id="index" title="出错啦"><p>ID无效.<br/>
		<%else

			if Request("del")="true" then
				set rs=server.CreateObject("adodb.recordset")
				Sql = "SELECT * FROM 74hu_gonggao where id="&id
				rs.open sql,conn,1,3
				if not (rs.bof and rs.eof)  then
				rs.delete
				end if
				Rs.close
				set rs=nothing%>
		<card id='index' title='删除公告'><p>
		已成功删除该公告!<br/>
<a href="gonggo.asp?sid=<%=sid%>">[公告管理]</a><br/>
			<%else%>
			<card id='index' title='删除公告'><p>
			是否要删除该公告?<br/>
			<a href='gonggo.asp?sid=<%=sid%>&amp;Action=del&amp;del=true&amp;id=<%=id%>'>[确定删除]</a><br/><anchor>返回上页<prev/></anchor><br/>
<a href="gonggo.asp?sid=<%=sid%>">[公告管理]</a><br/>
<%end if
end if
end Function
Function save
	name=Trim(Request("name"))
	title=Trim(Request("title"))
	id=TRim(Request("id"))
	if not isnumeric(id) then id=""
		if name=""  then errmsg=errmsg&"公告标题不能为空<br/>":flag=0
		if title=""  then errmsg=errmsg&"公告内容不能为空<br/>":flag=0
		if Request("edit")="edit" then
		if id=""  then errmsg=errmsg&"ID无效<br/>":flag=0
		if flag<>"0" then
				set rs=server.CreateObject("adodb.recordset")
				sql="select * from 74hu_gonggao where id="&id
				rs.open sql,conn,1,3
				rs("name")=name
				rs("title")=title
				rs.update()
				Rs.close
				set rs=nothing%>
<card id="index" title="修改公告" ontimer="gonggo.asp?sid=<%=sid%>"><timer value='10'/><p>
修改公告成功<br/><br/>
<a href="gonggo.asp?sid=<%=sid%>">[公告管理]</a><br/>
<%else
Response.Write("<card id='index' title='发布公告出错' ><p>") 
Response.Write(""&errmsg&"<br/><anchor>[返回修改]<prev/></anchor><br/>") 		
		end if
		else
			if flag<>"0" then
				set rs=server.CreateObject("adodb.recordset")
				sql="select * from 74hu_gonggao"
				rs.open sql,conn,1,3
				rs.addnew()
				rs("name")=name
				rs("title")=title
				rs("HU_time")=now()
				rs.update()
				Rs.close
				set rs=nothing%>
<card id="index" title="发布公告" ontimer="gonggo.asp?sid=<%=sid%>"><timer value='10'/>
<p>
发布公告成功<br/><br/>
<a href="gonggo.asp?sid=<%=sid%>">[公告管理]</a><br/>
<%else
Response.Write("<card id='index' title='发布公告出错' ><p>") 
Response.Write(""&errmsg&"<br/><anchor>[返回修改]<prev/></anchor><br/>") 	
		end if
		END IF
end Function%>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
</p></card></wml><%call CloseConn%>