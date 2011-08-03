<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<%Call Head()%>
<!--#include file="include.asp"-->
<card title="文章复制"><p>
<% dim TP
TP=request("TP")
if TP<>"" then
Dim Url,TextInfo
Url=CodeWML(Request.Form("Url"))
if len(url)< 10 then
 Response.Write "请正确输入文章地址！<br/>"
 Response.Write	"<anchor>返回上级<prev/></anchor> "
 Response.Write "</p></card></wml>"
 response.end
 End if
TextInfo=XMLHTTPGet(Url)
%>
文章地址:<br/>
<input name="url<%=minute(now)%><%=second(now)%>" value="<%=Url%>"/>
<br/>文章内容:<br/><input name="content<%=minute(now)%><%=second(now)%>" value="<%=CodeWml(RemoveHTML(TextInfo))%>"/><br/>
<%else%>
文章地址:<br/>
<input name="url<%=minute(now)%><%=second(now)%>" value="http://"/><br/>
<anchor>获取内容<go href="text.asp?sid=<%=sid%>" method="post">
<postfield name="url" value="$(url<%=minute(now)%><%=second(now)%>)"/><postfield name="TP" value="1"/></go></anchor><br/>
<%end if%>
----------<br/>
<%if TP<>"" then%>
<a href="text.asp?sid=<%=sid%>">[返回上级]</a><br/>
<%end if%>
<a href="index.asp?sid=<%=sid%>">[站长工具]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>