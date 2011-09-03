<!-- #include file="ding.asp" -->
<!--#include file="mymin.asp"-->
<%
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
IF KEY<>0 then
  Call Error("<card title=""出错""><p>你的权限不足！")
  end if

dim pp,id
pp=request.querystring("lx")
id=request.querystring("id")
if id="" or IsNumeric(id)=False then
  Call Error("<card title=""出错""><p>ID无效！")
  end if
if pp="" or IsNumeric(pp)=False then
  Call Error("<card title=""出错""><p>ID无效！")
  end if
dim wmlhead
wmlhead = "tools/wml.txt"

call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_class where classid="&id,conn,1,1

if rs.bof and rs.eof then
    	notclass
else
classright
end if
%>
<%sub classright()%>
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.2//EN" "http://www.wapforum.org/DTD/wml12.dtd">
<wml><%
if pp=10 then classname="最新文章"
if pp=11 then classname="最热文章"
if pp=12 then classname="随机文章"
if pp=19 then classname="站内搜框"
if pp=0 then classname="新的页面"
if pp=1 then classname="文章栏目"
if pp=8 then classname="随机广告"
if pp=9 then classname="WML标签"
if pp=20 then classname="WML页面"
%>
<%if pp=0 then %>
<card title="编辑页面菜单">
<p>名称:<%=noubb(rs("class"))%><br/>
        类型:<%=classname%><br/>
请输入新页面名称<br/><input name="class1<%=minute(now)%><%=second(now)%>" title="名称" value="<%=noubb(rs("class"))%>" emptyok="false"/><br/>
栏目后面:<select name="br<%=minute(now)%><%=second(now)%>" value="<%=noubb(rs("br"))%>">
			<option value="1">自动换行</option>
			<option value="2">不换行</option>
			</select><br/>
显示顺序:<input name="pid<%=minute(now)%><%=second(now)%>" type="text" value="<%=noubb(rs("pid"))%>" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="editclasssave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1<%=minute(now)%><%=second(now)%>)"/>
         <postfield name="pid" value="$(pid<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="br" value="$(br<%=minute(now)%><%=second(now)%>)"/>
        <postfield name='lx' value='0'/>
    </go>
</anchor>
<%elseif pp=1 then%>
<card title="编辑文章栏目">
<p>名称:<%=noubb(rs("class"))%><br/>
        类型:<%=classname%><br/>
请输入新栏目名称<br/><input name="class1<%=minute(now)%><%=second(now)%>" title="名称" value="<%=noubb(rs("class"))%>" emptyok="false"/><br/>
栏目后面:<select name="br<%=minute(now)%><%=second(now)%>" value="<%=noubb(rs("br"))%>">
			<option value="1">自动换行</option>
			<option value="2">不换行</option>
			</select><br/>
显示顺序:<input name="pid<%=minute(now)%><%=second(now)%>" type="text" value="<%=noubb(rs("pid"))%>" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="editclasssave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1<%=minute(now)%><%=second(now)%>)"/>
         <postfield name="pid" value="$(pid<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="br" value="$(br<%=minute(now)%><%=second(now)%>)"/>
        <postfield name='lx' value='1'/>
    </go>
</anchor>
<%elseif pp=2 then%>
<card title="编辑UBB标签">
<p>名称:<%=noubb(rs("class"))%><br/>
        类型:<%=classname%><br/>新UBB标签标记<br/><input name="class1<%=minute(now)%><%=second(now)%>" title="名称" value="<%=noubb(rs("class"))%>" emptyok="false"/><br/>
新UBB标签内容<br/><input name="wmltxt<%=minute(now)%><%=second(now)%>" title="名称" value="<%=noubb(rs("wmltxt"))%>" emptyok="false"/><br/>
栏目后面:<select name="br<%=minute(now)%><%=second(now)%>" value="<%=noubb(rs("br"))%>">
			<option value="1">自动换行</option>
			<option value="2">不换行</option>
			</select><br/>
显示顺序:<input name="pid<%=minute(now)%><%=second(now)%>" type="text" value="<%=noubb(rs("pid"))%>" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="editclasssave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="wmltxt" value="$(wmltxt<%=minute(now)%><%=second(now)%>)"/>
         <postfield name="pid" value="$(pid<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="br" value="$(br<%=minute(now)%><%=second(now)%>)"/>
        <postfield name='lx' value='2'/>
    </go>
</anchor>
<%elseif pp=9 then%>
<card title="修改WML标签">
<p>
        类型:<%=classname%><br/>新WML标签标记:<br/><input name="class1<%=minute(now)%><%=second(now)%>" title="名称" value="<%=noubb(rs("class"))%>" emptyok="false"/><br/>
新WML标签内容:<br/><input name="wmltxt<%=minute(now)%><%=second(now)%>" title="名称" value="<%=noubb(rs("wmltxt"))%>" emptyok="false"/><br/>
栏目后面:<select name="br<%=minute(now)%><%=second(now)%>" value="<%=noubb(rs("br"))%>">
			<option value="1">自动换行</option>
			<option value="2">不换行</option>
			</select><br/>
显示顺序:<input name="pid<%=minute(now)%><%=second(now)%>" type="text" value="<%=noubb(rs("pid"))%>" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="editclasssave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="wmltxt" value="$(wmltxt<%=minute(now)%><%=second(now)%>)"/>
         <postfield name="pid" value="$(pid<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="br" value="$(br<%=minute(now)%><%=second(now)%>)"/>
        <postfield name='lx' value='9'/>
    </go>
</anchor>
<%elseif pp=20 then%>
<card title="修改WML页面">
<p>
        类型:<%=classname%><br/>新WML页面标记:<br/><input name="class1<%=minute(now)%><%=second(now)%>" title="名称" value="<%=noubb(rs("class"))%>" emptyok="false"/><br/>
新WML页面内容:(无须文件头)<br/><input name="wmltxt<%=minute(now)%><%=second(now)%>" title="名称" value="<%=uubb(replace(LoadFile(rs("url")),LoadFile(wmlhead),""))%>" emptyok="false"/><br/>
栏目后面:<select name="br<%=minute(now)%><%=second(now)%>" value="<%=noubb(rs("br"))%>">
			<option value="1">自动换行</option>
			<option value="2">不换行</option>
			</select><br/>
显示顺序:<input name="pid<%=minute(now)%><%=second(now)%>" type="text" value="<%=noubb(rs("pid"))%>" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="wmlcl.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="wmltxt" value="$(wmltxt<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="url" value="<%=rs("url")%>"/>
         <postfield name="pid" value="$(pid<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="br" value="$(br<%=minute(now)%><%=second(now)%>)"/>
        <postfield name='lx' value='20'/>
    </go>
</anchor>

<%elseif pp=7 then%>
<card title="编辑文件栏目">
<p>名称:<%=noubb(rs("class"))%><br/>
        类型:<%=classname%><br/>新文件栏目名字:<br/><input name="class1<%=minute(now)%><%=second(now)%>" title="名称" value="<%=noubb(rs("class"))%>" emptyok="false"/><br/>
栏目后面:<select name="br<%=minute(now)%><%=second(now)%>" value="<%=noubb(rs("br"))%>">
			<option value="1">自动换行</option>
			<option value="2">不换行</option>
			</select><br/>
显示顺序:<input name="pid<%=minute(now)%><%=second(now)%>" type="text" value="<%=noubb(rs("pid"))%>" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="editclasssave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1<%=minute(now)%><%=second(now)%>)"/>
         <postfield name="pid" value="$(pid<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="br" value="$(br<%=minute(now)%><%=second(now)%>)"/>
        <postfield name='lx' value='7'/>
    </go>
</anchor>
<%elseif pp=8 then%>
<card title="编辑随机广告">
<p>名称:<%=noubb(rs("class"))%><br/>
        类型:<%=classname%><br/>新广告名称标记:<br/><input name="class1<%=minute(now)%><%=second(now)%>" title="名称" value="<%=noubb(rs("class"))%>"/><br/>
栏目后面:<select name="br<%=minute(now)%><%=second(now)%>" value="<%=noubb(rs("br"))%>">
			<option value="1">自动换行</option>
			<option value="2">不换行</option>
			</select><br/>
显示顺序:<input name="pid<%=minute(now)%><%=second(now)%>" type="text" value="<%=noubb(rs("pid"))%>" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="editclasssave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1<%=minute(now)%><%=second(now)%>)"/>
         <postfield name="URL" value="$(URL<%=minute(now)%><%=second(now)%>)"/>
         <postfield name="pid" value="$(pid<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="br" value="$(br<%=minute(now)%><%=second(now)%>)"/>
        <postfield name='lx' value='8'/>
    </go>
</anchor>
<%else%>
<card title="编辑调用栏目">
<p>名称:<%=noubb(rs("class"))%><br/>
        类型:<%=classname%><br/>新调用栏目名:<br/><input name="class1<%=minute(now)%><%=second(now)%>" title="名称" value="<%=noubb(rs("class"))%>" emptyok="false"/><br/>
调用栏目(留空则为全部栏目):<br/>栏目ID<input name="relid<%=minute(now)%><%=second(now)%>" title="条数" value="<%=noubb(rs("relid"))%>" emptyok="false"/><br/>
调用条数:<br/><input name="num<%=minute(now)%><%=second(now)%>" title="条数" value="<%=noubb(rs("num"))%>" emptyok="false"/><br/>
栏目后面:<select name="br<%=minute(now)%><%=second(now)%>" value="<%=noubb(rs("br"))%>">
			<option value="1">自动换行</option>
			<option value="2">不换行</option>
			</select><br/>
显示顺序:<input name="pid<%=minute(now)%><%=second(now)%>" type="text" value="<%=noubb(rs("pid"))%>" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="editclasssave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1<%=minute(now)%><%=second(now)%>)"/>
         <postfield name="pid" value="$(pid<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="br" value="$(br<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="num" value="$(num<%=minute(now)%><%=second(now)%>)"/>
        <postfield name="relid" value="$(relid<%=minute(now)%><%=second(now)%>)"/>
        <postfield name='lx' value='<%=pp%>'/>
    </go>
</anchor>
<%end if%>
<br/>----------<br/>
<% if rs("parent")<>0 then %>
<a href='Clist.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;lx=<%=pp%>'>[栏目分类]</a><br/>
<%end if%>
<a href='class.asp?sid=<%=sid%>'>[栏目分类]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml><%Response.End%><%End Sub%>
<%sub notclass()%><%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%>
<?xml version="1.0" encoding="UTF-8" ?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.3//EN" "http://www.wapforum.org/DTD/wml13.dtd">
<wml>
<head>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card title="出错了">
<p align="left">
没有这个栏目!
<br/>----------<br/>
<a href='class.asp?sid=<%=sid%>'>返回栏目管理</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml><%Response.End%><%End Sub%><%call CloseConn%>