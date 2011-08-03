<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<%Call Head()%>
<card title="网站攻击侦查"><p>
<%
IF KEY<>0 then
  response.write"你的权限不足！</p></card></wml>"
  response.end
  end if
act=request("act")
if act="sql" then
response.write"SQL注入攻击<br/>-------------<br/>一般要封锁IP，用于保护网站安全。因为想得出用攻击代码绝对不简单！特别是74hu_<br/>-------------<br/>"
call conndata
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from 74hu_sql order by id desc",conn,1,1
If Not rs.eof then
	PageSize=8
	gopage="sql.asp?act=sql&amp;sid="&sid&"&amp;"
	Count=conn.execute("Select count(ID) from 74hu_sql")(0)	
	page=int(request("page"))
	if page<=0 or page="" or isnumeric(page)=false then page=1
	pagecount=(count+pagesize-1)\pagesize
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For	
response.write ""&i+(page-1)*PageSize&".IP地址："&rs("HU_ip")&"<br/>记录时间："&fordate2(rs("HU_time"))&"<br/>非法字符："&rs("HU_str")&"<br/><br/>"
rs.moveNext
Next
	if page>0 then response.write "<br/>"
	if page>1 then response.write "<a href="""&gopage&"page=1"">首页</a>&nbsp;"
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>&nbsp;"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>&nbsp;"
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&pagecount&""">末页</a>"
	if pagecount>1 then response.write "<br/>第"&page&"页 共"&pagecount&"页<br/>第<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/>页 <a href="""&gopage&"page=$(page)"">跳转</a><br/>"
Else
response.write"暂时没有<br/> "
end if
rs.close
set rs=nothing
response.write "※网站管理中要经常查看攻击情况，以便及时防患，确保安全！注意同一IP的攻击情况，不怕贼偷窃，就怕贼惦记！<br/><a href='sql.asp?sid="&sid&"'>侦查后台</a>"
elseif act="dl" then
response.write"后台登陆攻击<br/>-------------<br/>一般不封锁IP，用于研究密码，制定人想不出的密码！<br/>-------------<br/>"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from 74hu_eyi order by id desc",conn,1,1
If Not rs.eof then
	PageSize=5
	gopage="sql.asp?act=dl&amp;sid="&sid&"&amp;"
	Count=conn.execute("Select count(ID) from 74hu_eyi")(0)	
	page=int(request("page"))
	if page<=0 or page="" or isnumeric(page)=false then page=1
	pagecount=(count+pagesize-1)\pagesize
	rs.move(pagesize*(page-1))
	For i=1 To PageSize
	If rs.eof Then Exit For	
response.write ""&i+(page-1)*PageSize&".IP地址："&rs("HU_ip")&"<br/>记录时间："&fordate2(rs("HU_time"))&"<br/>用户名："&rs("HU_name")&"<br/>密码："&rs("HU_pass1")&"<br/>用户名："&rs("HU_pass2")&"<br/><br/>"
rs.moveNext
Next
	if page>1 then response.write "<a href="""&gopage&"page=1"">首页</a>&nbsp;"
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&page+1&""">下页</a>&nbsp;"
	if page>1 then response.write "<a href="""&gopage&"page="&page-1&""">上页</a>&nbsp;"
if page-pagecount<0 then response.write "<a href="""&gopage&"page="&pagecount&""">末页</a>"
	if pagecount>1 then response.write "<br/>第"&page&"页 共"&pagecount&"页<br/>第<input name=""page"" format=""*N"" value="""&page&""" type=""text"" maxlength=""5"" emptyok=""true"" size=""3""/>页 <a href="""&gopage&"page=$(page)"">跳转</a><br/>"
Else
response.write"暂时没有<br/> "
end if

rs.close
set rs=nothing
response.write "※网站管理中要经常查看攻击情况，以便及时防患，确保安全！注意同一IP的攻击情况，不怕贼偷窃，就怕贼惦记！<br/><a href='sql.asp?sid="&sid&"'>侦查后台</a>"
else
response.write "<a href='sql.asp?act=sql&amp;sid="&sid&"'>SQL注入攻击</a><br/>"
response.write "<a href='sql.asp?act=dl&amp;sid="&sid&"'>后台登陆攻击</a><br/>-----------"
end if
Last
%>