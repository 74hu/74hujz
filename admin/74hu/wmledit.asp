<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<%Call Head()%>
<card title="编辑WML页面">
<p><%dim path,pathname,wmltxt,wmlhead,TP
wmlhead = "wml.txt"
path=trim(request("path"))
pathname=trim(request("pathname"))
TP=trim(request("TP"))
wmltxt=trim(request("wmltxt"))
if path="" then
  Call Error("WML页面地址无效！")
  end if
if pathname="" then
  Call Error("WML页面名称无效！")
  end if
function uubb(str)
	str=trim(str)
	if IsNull(str) then exit function
	str=replace(str,"&","&amp;")
	str=replace(str,"<","&lt;")
	str=replace(str,">","&gt;")
	str=replace(str,"'","&apos;")
	str=replace(str,"""","&quot;")
	uubb=str
end function
IF TP<>"" then
if wmltxt="" then
  Call Error("WML页面内容不能为空！")
  end if
        dim filename
        filename="/wml/"&pathname
	call SaveToFile(LoadFile(wmlhead)&wmltxt,filename)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filesize=fso.GetFile(Server.MapPath(filename)).size
response.write"WML页面编辑成功！<br/>"
else
%>
新WML页面内容:<br/><input name="wmltxt<%=minute(now)%><%=second(now)%>" title="名称" value="<%=uubb(replace(LoadFile(path),LoadFile(wmlhead),""))%>" emptyok="false"/><br/>
<anchor>确认提交
    <go href="wmledit.asp?sid=<%=sid%>" method="post" accept-charset="utf-8">
        <postfield name="wmltxt" value="$(wmltxt<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="path" value="<%=path%>"/>
        <postfield name="pathname" value="<%=pathname%>"/>
        <postfield name="TP" value="1"/>
    </go>
</anchor><br/>
<a href="wmledit2.asp?path=<%=path%>&amp;pathname=<%=pathname%>&amp;sid=<%=sid%>">2.0编辑WML</a><br/>
<a href="wmldel.asp?path=<%=path%>&amp;sid=<%=sid%>">删除此WML</a><br/>
<%end if%>
提示:WML页面应从&lt;card到&lt;/wml&gt;<br/>
<a href='wmltext.asp?sid=<%=sid%>'>[WML页面管理]</a>
<br/><a href="index.asp?sid=<%=sid%>">[站长工具]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card>
</wml>