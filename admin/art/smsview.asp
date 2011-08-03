
<!-- #include file="../ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="文章详情">
<p>
<% Dim id,ids,TP
TP=request("TP")
id=request("id")
if id="" or IsNumeric(id)=False then
  Call Error("ID无效！")
  end if
   ids=request("ids") 
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_article where id="&id,conn,1,1
	if rs.eof and rs.bof then
	response.write "文章不存在<br/>"
	else
Counts=rs("smspin")
Response.Write	"["&ubb(fordate(rs("HU_date")))&"]<br/>"
Response.Write	""&ubb(rs("title"))&"<br/>"
Content=rs("test")
if request.querystring("o")<>1 then
pageWordNum=viewtnums
StartWord = 1
Length=len(Content)
PageAll=(Length+PageWordNum-1)\PageWordNum
ii=clng(request("ii"))
i=clng(request("i"))
if ii<>0 then i=ii-1
if isnull(i) or i="" then i=0

	dim ccc,sss
	ccc=instr(content,"||")
	if ccc>0 then
	sss=split(content,"||")
	PageAll=ubound(sss)+1
		if i>PageAll-1 then i=PageAll-1
	content = sss(i)

	else
		if clng(i)>int(PageAll) then i=PageAll-1
	Content = mid(Content,StartWord+i*PageWordNum,PageWordNum)
	end if
response.write("-----------<br/>" &ubb(content)& "")
if 0<=i<PageAll then
       Response.Write "<br/>"
end if
   if cint(i)<cint(PageAll)-1 then
       Response.Write "<a href='smsview.asp?ids="&ids&"&amp;id=" &  rs("id") & "&amp;i=" & i+1 & "&amp;p=" & p & "&amp;TP="&TP&"&amp;sid="&sid&"'>下页</a>"
   End if
   if cint(i)>0 then 
   Response.Write "&nbsp;" & "<a href='smsview.asp?ids="&ids&"&amp;id=" &  rs("id") & "&amp;i=" & i-1 & "&amp;p=" & p & "&amp;TP="&TP&"&amp;sid="&sid&"'>上页</a>"
		if  i<pageall-1 then Response.Write "&nbsp;<a href='smsview.asp?ids="&ids&"&amp;id=" &  rs("id") & "&amp;i=100&amp;p=" & p & "&amp;TP="&TP&"&amp;sid="&sid&"'>尾页</a>"
   End if
if PageAll>1 then
response.write ("&nbsp;<a href='smsview.asp?id=" &  rs("id") & "&amp;ids="&ids&"&amp;p="&p&"&amp;o=1&amp;TP="&TP&"&amp;sid="&sid&"'>全文</a>")
       response.write "(" & i+1 & "/" & PageAll & ")"
%>
<br/>第<input name="i<%=minute(now)%><%=second(now)%>" title="页码" type="text" format="*N" emptyok="true" size="2" value="<%response.write(i+2)%>" maxlength="2"/>
<anchor>跳页
    <go href="smsview.asp?id=<%=id%>&amp;ids=<%=ids%>&amp;p=<%=p%>&amp;TP=<%=TP%>&amp;sid=<%=sid%>" accept-charset='utf-8'>
        <postfield name="ii" value="$(i<%=minute(now)%><%=second(now)%>)"/>
    </go>
</anchor><br/>
<%
end if
else
response.write("-----------<br/>" & ubb(content) & "")
response.write ("<br/><a href='smsview.asp?id=" &  rs("id") & "&amp;ids="&ids&"&amp;p="&p&"&amp;TP="&TP&"&amp;sid="&sid&"'>分页显示</a><br/>")
end if 
end if 
%>
<br/>----------<br/>
<a href="wzgl.asp?sid=<%=sid%>&amp;id=<%=rs("id")%>&amp;classid=<%=ids%>">[文章管理]</a><br/>
<a href="adminsmscl.asp?sid=<%=sid%>&amp;id=<%=ids%>">[文章列表]</a>
<br/><a href="wzclass.asp?sid=<%=sid%>">[返回分类]</a><br/>
<%if TP<>"" then%>
<a href="wzcl.asp?sid=<%=sid%>">[文章管理]</a><br/>
<%end if%>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing%>