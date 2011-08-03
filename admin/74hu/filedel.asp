<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<%Call Head()%>
<card id="index" title="文件删除"><p>
<%
dim path,TP,ok
path=trim(request.querystring("path"))
ok=trim(request.querystring("ok"))
TP=trim(request.querystring("TP"))
if path="" then
  Call Error("文件地址无效！")
  end if
if ok<>"" then
path=Server.MapPath(path)
Set fso = server.CreateObject("Scripting.FileSystemObject")
	IF (fso.FileExists(path)) Then
	Set delfile = fso.GetFile(path)
	delfile.Delete
	set delfile=nothing
	response.write("文件删除成功！<br/>")
	else
	response.write(path & "文件不存在！<br/>")
	end if
set fso=nothing
else
Response.write "您确认要删除该文件？<br/>"
Response.write "<a href=""filedel.asp?path="&path&"&amp;TP="&TP&"&amp;ok=1&amp;sid="&sid&""">是的</a><br/>"
Response.write "<anchor>取消<prev/></anchor><br/>"
end if
%><a href='fileman.asp?TP=<%=TP%>&amp;sid=<%=sid%>'>[返回上级]</a><br/>
<a href="Files.asp?sid=<%=sid%>">[文件管理]</a><br/>
<a href="index.asp?sid=<%=sid%>">[站长工具]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>