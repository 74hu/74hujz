<!--#include file="../ding.asp"-->
<!-- #include file="../mymin.asp" -->
<!-- #include file="upload.inc" -->
<%Call Head()%>
<card title="新建WML页面">
<p><%dim wmlname,wmltxt,wmlhead,TP
wmlhead = "wml.txt"
wmlname=trim(request("wmlname"))
TP=trim(request("TP"))
wmltxt=trim(request("wmltxt"))
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
     if Wmlname="" then
        filename="/wml/"&addwml(now())
        else
        filename="/wml/"&Wmlname&".wml"
        end if
	call SaveToFile(LoadFile(wmlhead)&wmltxt,filename)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filesize=fso.GetFile(Server.MapPath(filename)).size
response.write"WML页面添加成功！<br/>"
response.write"预览:<a href='"&filename&"'>"&filename&"</a><br/>"
else
%>
WML页面内容:(无须文件头)<br/><input name="wmltxt<%=minute(now)%><%=second(now)%>" title="WML页面内容" value="" emptyok="false"/><br/>
WML地址命名:(留空则自动命名)<br/><input name="wmlname<%=minute(now)%><%=second(now)%>" title="WML地址命名" value="" emptyok="true"/><br/>
<anchor>确认提交
    <go href="wmladd.asp?sid=<%=sid%>" method="post" accept-charset="utf-8">
        <postfield name="wmltxt" value="$(wmltxt<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="wmlname" value="$(wmlname<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="TP" value="1"/><br/>
    </go>
</anchor>
提示:WML页面应从&lt;card&gt;到&lt;/wml&gt;<br/>
<%end if%>
<a href='wmltext.asp?sid=<%=sid%>'>[WML页面管理]</a>
<br/><a href="index.asp?sid=<%=sid%>">[站长工具]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card>
</wml>