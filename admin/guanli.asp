<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<!-- #include file="md5.asp" -->
<%Call Head()%>
<%
dim siddd
Function SID_RANDOMIZE()
	Dim c,id
	randomize 
	c="a,A,1,b,B,2,c,C,3,d,D,4,e,E,5,f,F,6,g,G,7,h,H,8,i,I,9,j,J,0,k,K,1,l,L,2,m,M,3,n,N,4,o,O,5,p,P,6,q,Q,7,r,R,8,s,S,9,t,T,0,u,v,w,x,y,z,"
	id=split(c,",") 
	for i=1 to 10
	SID_RANDOMIZE=int(rnd()*36)&SID_RANDOMIZE
	Next
	SID_RANDOMIZE=SID_RANDOMIZE&minute(now)&second(now)&day(now)&month(now)
End Function

siddd=md5(md5(sjhm+SID_RANDOMIZE,16),32)
IF  Request.QueryString("Action")="view" Then
	if KEYid<>1 then 
		Call Error("<card title=""警告""><p>权限不足！")
	end if
	call view
elseiF  Request.QueryString("Action")="del" Then
	if KEYid<>1 then 
		Call Error("<card title=""警告""><p>权限不足！")
	end if
	call del
elseiF  Request.QueryString("Action")="edit" Then
	if KEYid<>1 then 
		Call Error("<card title=""警告""><p>权限不足！")
	end if
	call edit
elseiF  Request.QueryString("Action")="add" Then
	if KEYid<>1 then 
		Call Error("<card title=""警告""><p>权限不足！")
	end if
	call add
elseiF  Request.QueryString("Action")="save" Then
	if KEYid<>1 then 
		Call Error("<card title=""警告""><p>权限不足！")
	end if
	call save
else
	call index
end if

Function index%>
	<card id="index" title="设定管理员">
	<p align="left">
	<%
	Set Rs = Server.CreateObject("Adodb.Recordset")
	Sql = "SELECT * FROM 74hu_admin order by id asc"
	Rs.Open Sql,conn,1,1
	If Not rs.eof Then
	PageSize=10
	gopage="guanli.asp?sid="&sid&"&amp;"
	Count=rs.recordcount
	page=int(request.QueryString ("page"))
	if page<=0 or page="" then page=1
	pagecount=(count+pagesize-1)\pagesize
	if page>pagecount then page=pagecount
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For	
	if rs("key")=2 then keyname="普通管理员"
	if rs("key")=0 then keyname="超级管理员"
	%><%=i+(page-1)*PageSize%>.<a href='guanli.asp?sid=<%=sid%>&amp;Action=view&amp;id=<%=noubb(rs("id"))%>'><%=keyname%><%=noubb(rs("username"))%><%if rs("id")=1 then%>站长<%end if%></a><br/>
	上次管理:<%=noubb(fordate(rs("lastdate")))%><br/>
	<%     
	rs.moveNext
	Next
	if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>"
	if pagecount>1 then response.write "<br/><b>"&page&"</b>/"&pagecount&"页<input username=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/><a href="""&gopage&"page=$(page)"">>></a>"
	Else
	%>
	(没有管理)
	<%end if%>
	<%
	Rs.close
	set rs=nothing
	Response.Write("<br/><a href='guanli.asp?sid="&sid&"&amp;Action=add'>[添加管理]</a><br/>")
end Function
Function add%>
	<card id="add" title="添加管理"><p>
	用户:<br/><input name="username" maxlength="255" /><br/>
	密码:<br/><input name="password" /><br/> 
	高密:<br/><input name="password2" /><br/> 
	管理范围:<br/><select name="key">
	<option value="0">超级管理员</option>
	<option value="2">普通管理员</option>
	</select><br/>
	<anchor>添加
	<go href="guanli.asp?sid=<%=sid%>&amp;Action=save&amp;edit=add" method="post">
	<postfield name="username" value="$(username:n)" />
	<postfield name="password" value="$(password:n)" />
	<postfield name="password2" value="$(password2:n)" />
	<postfield name="key" value="$(key:n)" />
	</go>
	</anchor><br/>
	<a href="guanli.asp?sid=<%=sid%>">[设定管理]</a><br/>
	<%
end Function
Function edit
	id=TRim(Request("id"))
	if not isnumeric(id) then id=""
	if id=""  then%>
	<card id="index" title="出错啦"><p>ID无效.<br/>
	<%else
	Set Rs = Server.CreateObject("Adodb.Recordset")
	Sql = "SELECT * FROM 74hu_admin where id="&id
	Rs.Open Sql,conn,1,1
	if not (rs.bof and rs.eof)  then %>
	<card id="add" title="修改管理"><p>
	用户:<br/><input name="username" value="<%=rs("username")%>" maxlength="255" /><br/>
	密码:<br/><input name="password" value="" /><br/> 
	高级密码:<br/><input name="password2" value="" /><br/> 
	<%if id<>1 then%>
	管理范围:<br/><select name="key" value="<%=rs("key")%>">
	<option value="0">超级管理员</option>
	<option value="2">普通管理员</option>
	</select>
	<%end if%>
	<br/><anchor>修改
	<go href="guanli.asp?sid=<%=sid%>&amp;Action=save&amp;edit=edit" method="post">
	<postfield name="username" value="$(username:n)" />
	<postfield name="id" value="<%=id%>" />
	<postfield name="password" value="$(password:n)" />
	<postfield name="password2" value="$(password2:n)" />
	<%if id<>1 then%>	
	<postfield name="key" value="$(key:n)" />
	<%end if%>
	</go>
	</anchor><br/>
	提示:修改管理员将会提示重新登录.<br/>
	<a href="guanli.asp?sid=<%=sid%>">[设定管理]</a><br/>
	<%else%>
	<card id="index" title="出错啦"><p>没有该管理!<br/>
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
	<card id="index" title="设定管理">
	<p>
	<%
	Set Rs = Server.CreateObject("Adodb.Recordset")
	Sql = "SELECT * FROM 74hu_admin where id="&id
	Rs.Open Sql,conn,1,1
	if not (rs.bof and rs.eof)  then %>
	用户:<%=noubb(rs("username"))%><br/>
	密码:<%=noubb(rs("password"))%><br/>
	高密:<%=noubb(rs("HU_admin"))%><br/>
	上次管理:<%=noubb(fordate(rs("lastdate")))%><br/>
	上次IP:<%=noubb(rs("lastip"))%><br/>
	<a href='guanli.asp?sid=<%=sid%>&amp;Action=edit&amp;id=<%=id%>'>[编辑]</a> <a href='guanli.asp?sid=<%=sid%>&amp;Action=del&amp;id=<%=id%>'>[删除]</a><br/>
	<a href="guanli.asp?sid=<%=sid%>">[设定管理]</a><br/>
	<%else%>
	暂无管理!<br/>
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
	Sql = "SELECT * FROM 74hu_admin where id="&id
	rs.open sql,conn,1,3
	if not (rs.bof and rs.eof)  then
	if id=1 then
	Call Error("<card title='删除管理员'><p>站长不能删除！")
	end if
	rs.delete
	end if
	Rs.close
	set rs=nothing%>
	<card id='index' title='删除管理员'><p>
	已成功删除该管理员!<br/>
	<a href="guanli.asp?sid=<%=sid%>">[设定管理]</a><br/>
	<%else%>
	<card id='index' title='删除管理员'><p>
	是否要删除该管理员?<br/>
	<a href='guanli.asp?sid=<%=sid%>&amp;Action=del&amp;del=true&amp;id=<%=id%>'>[确定删除]</a><br/><anchor>返回上页<prev/></anchor><br/>
	<a href="guanli.asp?sid=<%=sid%>">[设定管理]</a><br/>
	<%end if
	end if
end Function
Function save
	username=Trim(Request("username"))
	password=Trim(Request("password"))
	password2=Trim(Request("password2"))
	key=Trim(Request("key"))
	id=clng(Request("id"))
	if id<>1 then
	if key="" or isnumeric(key) =false then
	Call Error("管理范围无效！")
	end if
	end if
	if username=""  then errmsg=errmsg&"管理名称不能为空<br/>":flag=0
	if password=""  then errmsg=errmsg&"管理密码不能为空<br/>":flag=0
	if password2=""  then errmsg=errmsg&"管理高级密码不能为空<br/>":flag=0
	if Request("edit")="edit" then
	if id=""  then errmsg=errmsg&"ID无效<br/>":flag=0
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select ID from 74hu_admin where username='"&username&"' and id<>"&id,conn,1,1
	if not rs.eof then
	errmsg=errmsg&"管理员用户名已被使用<br/>":flag=0
	end if
	rs.close
	set rs=nothing
	if flag<>"0" then
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from 74hu_admin where id="&id
	rs.open sql,conn,1,3
	if rs.eof then
	errmsg=errmsg&"ID无效<br/>":flag=0
	end if
	rs("username")=username
	rs("password")=md5(md5(password,16),32)
	rs("HU_admin")=md5(md5(password2,16),32)
	if ID<>1 then
	rs("key")=key
	end if
	rs("sid")=setFilter(siddd)
	rs.update()
	Rs.close
	set rs=nothing%>
	<card id="index" title="修改管理" ontimer="guanli.asp?sid=<%=sid%>"><timer value='10'/><p>
	修改管理成功<br/><br/>
	<a href="guanli.asp?sid=<%=sid%>">[设定管理]</a><br/>
	<%else
	Response.Write("<card id='index' title='添加管理出错' ><p>") 
	Response.Write(""&errmsg&"<br/><anchor>[返回修改]<prev/></anchor><br/>")
	end if
	else
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open"select ID from 74hu_admin where username='"&username&"' and id<>"&id,conn,1,1
	if not rs.eof then
	errmsg=errmsg&"管理员用户名已被使用<br/>":flag=0
	end if
	rs.close
	set rs=nothing
	if flag<>"0" then
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from 74hu_admin"
	rs.open sql,conn,1,3
	rs.addnew()
	rs("username")=username
	rs("password")=md5(md5(password,16),32)
	rs("HU_admin")=md5(md5(password2,16),32)
	rs("sid")=siddd
	rs("key")=key
	rs.update()
	Rs.close
	set rs=nothing%>
	<card id="index" title="添加管理" ontimer="guanli.asp?sid=<%=sid%>"><timer value='10'/>
	<p>
	添加管理成功<br/><br/>
	<a href="guanli.asp?sid=<%=sid%>">[设定管理]</a><br/>
	<%else
	Response.Write("<card id='index' title='添加管理出错' ><p>") 
	Response.Write(""&errmsg&"<br/><anchor>[返回修改]<prev/></anchor><br/>") 	
	end if
	END IF
end Function%>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card>
</wml><%call CloseConn%>