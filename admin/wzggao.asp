<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<%     IF KEY<>0 then
  Call Error("<card title=""出错""><p>你的权限不足！")
  end if

        TP=Request.QueryString("TP")
        if TP="" or Isnumeric(TP)=False then TP=1
        TP=abs(TP)
        if TP=1 then YD="主站页面广告"
        if TP=2 then YD="推荐栏目广告"
        if TP=3 then YD="底部栏目广告"
        if TP=4 then YD="开发广告备用"
        if TP=5 then YD="开发广告备用"
	IF  (Request.QueryString("Action")="view") Then
		call view
	elseIF  (Request.QueryString("Action")="del") Then
		call del
	elseIF  (Request.QueryString("Action")="edit") Then
		call edit
	elseIF  (Request.QueryString("Action")="save") Then
		call save
	elseIF  (Request.QueryString("Action")="add") Then
		call add
	else
		call login
	End IF
%>
<%sub login%>
<card id="login" title="<%=YD%>"><p>
<%
call conndata
		 Set Rs = Server.CreateObject("Adodb.Recordset")
		Sql = "SELECT * FROM 74hu_gogo where typeID="&TP
		Rs.Open Sql,conn,1,1
If Not rs.eof	Then
	Dim PageSize,i
	PageSize=10							
	Dim Count,page,pagecount,gopage,parent			
	gopage="wzggao.asp?sid="&sid&"&amp;TP="&TP&"&amp;"
	Count=rs.recordcount
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1		
	pagecount=(count+pagesize-1)\pagesize	
        if page>pagecount then page=pagecount	
	rs.move(pagesize*(page-1))				
	For i=1 To PageSize       						
	If rs.eof Then Exit For				
%><%=i+(page-1)*PageSize%>.<a href="wzggao.asp?sid=<%=sid%>&amp;Action=view&amp;id=<%=Rs("id")%>&amp;TP=<%=TP%>"><%=ubb(Rs("name"))%></a><br/>点击：<%=Rs("tid")%><br/><%     
	rs.moveNext								
	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a>"
Else
%>
	(没有广告)

<%end if%><br/>---------<br/>
<a href="wzggao.asp?sid=<%=sid%>&amp;Action=add&amp;TP=<%=TP%>">[添加广告]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a></p></card>
<%end sub%>
<%sub del
call conndata
id=Request("id")
if not isnumeric(id) then id=""
if id<>"" then
	IF  (Request.QueryString("del")="true") Then
	Set Rs2 = Server.CreateObject("Adodb.Recordset")
	Sql2 = "select * FROM 74hu_gogo WHERE id=" & id &" and typeID="&TP
	Rs2.Open Sql2,conn,1,3
			if not (rs2.eof and rs2.bof) then
			rs2.Delete
			end if
			rs2.Close
			Set rs2 = Nothing%>
			<card id="revert" title="删除广告">
			<p>
			删除广告成功!<br/>
<a href="wzggao.asp?sid=<%=sid%>&amp;TP=<%=TP%>">[广告管理]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
			<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
			</p>
			</card>
	<%else%>
	<card id="revert" title="删除广告">
	<p>
	注意：确定要删除该广告吗？<br/>
<a href="wzggao.asp?Action=del&amp;id=<%=id%>&amp;del=true&amp;sid=<%=sid%>&amp;TP=<%=TP%>">[删除]</a> <a href="wzggao.asp?Action=login&amp;id=<%=id%>&amp;sid=<%=sid%>&amp;TP=<%=TP%>">[取消]</a><br/>
	<a href="wzggao.asp?sid=<%=sid%>&amp;TP=<%=TP%>">[广告管理]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
	<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
	</p></card>
	<%End IF
else%>
	<card id="revert" title="删除广告">
	<p>
	请不要非法传递参数!<br/>
	<a href="wzggao.asp?sid=<%=sid%>&amp;TP=<%=TP%>">[广告管理]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
		<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
		</p>
		</card>
<%end if
end sub%>
<%sub view
call conndata
id=Request("id")
if not isnumeric(id) then id=""
if id<>"" then
	Set Rs2 = Server.CreateObject("Adodb.Recordset")
	Sql2 = "select * FROM 74hu_gogo WHERE id=" & id &" and typeID="&TP
	Rs2.Open Sql2,conn,1,1
		if not (rs2.eof and rs2.bof) then%>
		<card id="revert" title="查看广告">
		<p>
        广告名称:<%=rs2("name")%><br/>
        点击次数:<%=rs2("tid")%><br/>
	广告地址:<%=ubb(rs2("url"))%><br/><br/>
<a href="wzggao.asp?Action=edit&amp;id=<%=Rs2("id")%>&amp;sid=<%=sid%>&amp;TP=<%=TP%>">[修改]</a> <a href="wzggao.asp?Action=del&amp;id=<%=Rs2("id")%>&amp;sid=<%=sid%>&amp;TP=<%=TP%>">[删除]</a><br/>
<a href="wzggao.asp?sid=<%=sid%>&amp;TP=<%=TP%>">[广告管理]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
		<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
		</p></card>
<%
		end if
		rs2.Close
		Set rs2 = Nothing

else
%>
	<card id="revert" title="查看广告">
	<p>
	请不要非法传递参数!<br/>
	<a href="wzggao.asp?sid=<%=sid%>&amp;TP=<%=TP%>">[广告管理]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
		<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
		</p>
		</card>
<%
end if
end sub%>
<%sub add
call conndata
	if Request.QueryString("add")="true"  then
		Name=Request.form("Name")
		url=Request.form("url")
		if name<>"" and url<>"" then
			set rs=Server.CreateObject("Adodb.Recordset")
			rs.open "select * from 74hu_gogo",conn,1,3
			rs.addnew()	
			rs("Name")=Name
			rs("typeID")=TP
			rs("url")=url
			rs.update
			rs.close
			set rs=nothing%>
			<card id="login" title="添加广告">
			<p>
			添加广告成功！	<br/>	
<a href="wzggao.asp?sid=<%=sid%>&amp;TP=<%=TP%>">[广告管理]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
		<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
		</p>
		</card>
		<%else%>
		<card id="login" title="添加广告">
		<p>
		广告名称或广告地址不能为空!<br/>
		<br/><a href="wzggao.asp?Action=add&amp;sid=<%=sid%>&amp;TP=<%=TP%>">返回修改</a><br/>
	<a href="wzggao.asp?sid=<%=sid%>&amp;TP=<%=TP%>">[广告管理]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
			<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
			</p>
			</card>
		<%end if
	else%> 
		<card id="login" title="添加广告">
		<p>
		广告名称:<br/>
		<input name="name" emptyok="false" maxlength="50"/><br/>
		广告地址:<br/>
		<input name="url" emptyok="false" maxlength="255"/><br/>
		<anchor>[添加广告]
			<go href="wzggao.asp?Action=add&amp;add=true&amp;sid=<%=sid%>&amp;TP=<%=TP%>" method="post">
				<postfield name="url" value="$(url)" />
				<postfield name="name" value="$(name)" />
				<postfield name="Method" value="pass" />
			</go>
		</anchor><br/>
<a href="wzggao.asp?sid=<%=sid%>&amp;TP=<%=TP%>">[广告管理]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
	<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
	</p></card>
	<%end if

end sub%>
<%sub edit
call conndata
id=TRim(Request("id"))
if not isnumeric(id) then id=""
if id<>"" then
if Request.QueryString("add")="true"  then
Name=Request.form("Name")
url=Request.form("url")
if name<>"" and url<>"" then
set rs=Server.CreateObject("Adodb.Recordset")
rs.open "select * from 74hu_gogo where id= "&id&" and typeID="&TP,conn,1,3
			rs("Name")=Name
			rs("url")=url
			rs.update
			rs.close
			set rs=nothing%>
			<card id="login" title="修改广告">
			<p>
			修改广告成功！<br/>
<a href="wzggao.asp?sid=<%=sid%>&amp;TP=<%=TP%>">[广告管理]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
			<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
			</p>
			</card>

		<%else%>
		<card id="login" title="修改广告">
		<p>
		广告名或广告联接不能为空!<br/>
		<br/><a href="wzggao.asp?Action=edit&amp;id=<%=id%>&amp;sid=<%=sid%>&amp;TP=<%=TP%>">返回修改</a><br/>
<a href="wzggao.asp?sid=<%=sid%>&amp;TP=<%=TP%>">[广告管理]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
		<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
		</p>
		</card>

	<%end if

	else
	Set Rs2 = Server.CreateObject("Adodb.Recordset")
	Sql2 = "select * FROM 74hu_gogo WHERE id=" & id &" and typeID="&TP
	Rs2.Open Sql2,conn,1,1
	if not (rs2.eof and rs2.bof) then%> 

	<card id="login" title="修改广告">
	<p>
	广告名称:<br/>
	<input name="name<%=tt%>" emptyok="false" maxlength="50" value="<%=rs2("name")%>"/><br/>
	广告地址:<br/>
	<input name="url<%=tt%>" emptyok="false" maxlength="255" value="<%=ubb(rs2("url"))%>"/><br/>
<anchor>修改广告<go href="wzggao.asp?Action=edit&amp;add=true&amp;id=<%=id%>&amp;sid=<%=sid%>&amp;TP=<%=TP%>" method="post">
<postfield name="url" value="$(url<%=tt%>)" />
<postfield name="name" value="$(name<%=tt%>)" />
</go>
</anchor><br/>
<a href="wzggao.asp?sid=<%=sid%>&amp;TP=<%=TP%>">[广告管理]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
</p></card>
<%
	end if
	rs2.Close
	Set rs2 = Nothing
end if
else
%>
<card id="revert" title="查看广告">
<p>出错了<br/>
<a href="wzggao.asp?sid=<%=sid%>&amp;TP=<%=TP%>">[广告管理]</a><br/>
<a href="ggao.asp?sid=<%=sid%>">[广告中心]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a><br/>
</p></card>
<%
end if
end sub
call CloseConn%>
</wml>