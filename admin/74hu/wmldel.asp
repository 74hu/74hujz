<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<%Call Head()%>
<card id="index" title="WAP页面删除"><p>
<%
dim path,TP
path=trim(request.querystring("path"))
TP=trim(request.querystring("TP"))
if path="" then
  Call Error("WML页面地址无效！")
  end if
if TP<>"" then
path=Server.MapPath(path)
Set fso = server.CreateObject("Scripting.FileSystemObject")
	IF (fso.FileExists(path)) Then
	Set delfile = fso.GetFile(path)
	delfile.Delete
	set delfile=nothing
	response.write("WML页面删除成功！<br/>")
	else
	response.write(path & "WML页面不存在！<br/>")
	end if
set fso=nothing
else
Response.write "您确认要删除该WML页面？<br/>"
Response.write "<a href=""wmldel.asp?path="&path&"&amp;TP=1&amp;sid="&sid&""">是的</a><br/>"
Response.write "<anchor>取消<prev/></anchor><br/>"
end if
%><a href='wmltext.asp?sid=<%=sid%>'>[WML页面]</a><br/>
<a href="index.asp?sid=<%=sid%>">[站长工具]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>