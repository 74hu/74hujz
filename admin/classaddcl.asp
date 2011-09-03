<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<%
IF KEY<>0 then
  Call Error("<card title=""出错""><p>你的权限不足！")
  end if
dim pp,id,idd,lxl
id=request.QueryString("id")
 idd= request.QueryString("idd")
 lxl= request.QueryString("lxl")
if id="" then id=0
pp=request.QueryString("pp")
if pp="" then pp=1
%>
<%if pp=0 then%>
<card title="添加新的页面">
<p>输入页面名称:<br/><input name="class1" title="名称" emptyok="false"/><br/>
栏目后面:<select name="br" value="2"><option value="1">自动换行</option><option value="2">不换行</option></select><br/>
显示顺序:<input name="pid" type="text" value="5" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="classsave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1)"/>
         <postfield name="pid" value="$(pid)"/>
        <postfield name="br" value="$(br)"/>
        <postfield name='lx' value='0'/>
        <postfield name='parent' value='<%=id%>'/>
    </go>
</anchor>
<%elseif pp=1 then%>
<card title="添加文章栏目">
<p>文章栏目名称:<br/>
<%
call conndata
set rs=server.createobject("adodb.recordset")
	sql="select * from 74hu_list"
	rs.open sql,conn,1,1
%><input name="class1" title="名称" value=""/><br/>
选择类别:
<select name="lid">
<% do while not rs.eof
%> <option value='<%=rs("classid")%>'><%=rs("class")%></option>   
<%  rs.movenext
        loop
%></select><br/>
栏目后面:<select name="br" value="1"><option value="1">自动换行</option><option value="2">不换行</option></select><br/>
显示顺序:<input name="pid" type="text" value="5" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="classsave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1)"/>
         <postfield name="pid" value="$(pid)"/>
        <postfield name="br" value="$(br)"/>
        <postfield name="lid" value="$(lid)"/>
        <postfield name='lx' value='1'/>
        <postfield name='parent' value='<%=id%>'/>
    </go>
</anchor>
<%elseif pp=2 then%>
<card title="添加UBB标签">
<p><a href="ubbcl.asp?sid=<%=sid%>">[UBB说明]</a><br/>
UBB后台标记:<br/><input name="class1" title="名称" emptyok="false"/><br/>
UBB内容:<br/><input name="wmltxt" title="名称" emptyok="false"/><br/>
栏目后面:<select name="br" value="1">
<option value="1">自动换行</option>
<option value="2">不换行</option>
</select><br/>
显示顺序:<input name="pid" type="text" value="5" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="classsave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1)"/>
        <postfield name="wmltxt" value="$(wmltxt)"/>
         <postfield name="pid" value="$(pid)"/>
        <postfield name="br" value="$(br)"/>
        <postfield name='lx' value='2'/>
        <postfield name='parent' value='<%=id%>'/>
    </go>
</anchor>
<%elseif pp=7 then%>
<card title="添加调用栏目">
<p>调用栏目名称:<br/><input name="class1" title="名称" emptyok="false"/><br/>
选择调用类型<br/><select name="lx">
<option value="10">最新文章</option>
<option value="11">最热文章</option>
<option value="12">随机文章</option>
<option value="19">站内搜框</option>
</select><br/>
调用栏目/论坛(留空则为全部栏目/论坛):<br/>栏目ID<input name="relid" title="栏目ID" emptyok="false" format="*N"/><br/>
调用条数:(站内搜框除外)<br/><input name="num" title="条数" emptyok="false" format="*N"/><br/>
栏目后面:<select name="br" value="1"><option value="1">自动换行</option><option value="2">不换行</option></select><br/>
显示顺序:<input name="pid" type="text" value="5" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="classsave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1)"/>
         <postfield name="img" value="$(img)"/>
         <postfield name="pid" value="$(pid)"/>
        <postfield name="br" value="$(br)"/>
        <postfield name='lx' value='$(lx)'/>
         <postfield name="num" value="$(num)"/>
         <postfield name="relid" value="$(relid)"/>
        <postfield name='parent' value='<%=id%>'/>
    </go>
</anchor>
<%elseif pp=8 then%>
<card title="添加随机广告">
<p>广告标题(用于后台标识):<br/><input name="class1" title="名称" emptyok="false"/><br/>
栏目后面:<select name="br" value="1"><option value="1">自动换行</option><option value="2">不换行</option></select><br/>
显示顺序:<input name="pid" type="text" value="5" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="classsave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1)"/>
         <postfield name="pid" value="$(pid)"/>
        <postfield name="br" value="$(br)"/>
        <postfield name='lx' value='8'/>
        <postfield name='parent' value='<%=id%>'/>
    </go>
</anchor>
<%elseif pp=9 then%>
<card title="添加WML标签">
<p><a href="faq.asp?p=14&amp;sid=<%=sid%>">[WML说明]</a><br/>
WML后台标记:<br/><input name="class1" title="名称" emptyok="false"/><br/>
WML标签内容:<br/><input name="wmltxt" title="名称" emptyok="false"/><br/>
栏目后面:<select name="br" value="1"><option value="1">自动换行</option><option value="2">不换行</option></select><br/>
显示顺序:<input name="pid" type="text" value="5" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="classsave.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1)"/>
        <postfield name="wmltxt" value="$(wmltxt)"/>
         <postfield name="pid" value="$(pid)"/>
        <postfield name="br" value="$(br)"/>
        <postfield name='lx' value='9'/>
        <postfield name='parent' value='<%=id%>'/>
    </go>
</anchor>
<%elseif pp=10 then%>
<card title="添加WML页面">
<p>WML页面后台标记:(无须文件头)<br/><input name="class1" title="名称" emptyok="false"/><br/>
WML地址命名:(留空则自动命名)<br/><input name="Wmlname" title="名称" emptyok="true"/><br/>
WML页面内容:<br/><input name="wmltxt" title="名称" emptyok="false"/><br/>
栏目后面:<select name="br" value="1"><option value="1">自动换行</option><option value="2">不换行</option></select><br/>
显示顺序:<input name="pid" type="text" value="5" format="*N" size="2" emptyok="false"/><br/>
<anchor>确认提交
    <go href="wmlcl.asp?sid=<%=sid%>&amp;id=<%=id%>" method="post" accept-charset="utf-8">
        <postfield name="class1" value="$(class1)"/>
        <postfield name="Wmlname" value="$(Wmlname)"/>
        <postfield name="wmltxt" value="$(wmltxt)"/>
         <postfield name="pid" value="$(pid)"/>
        <postfield name="br" value="$(br)"/>
        <postfield name='lx' value='20'/>
        <postfield name='parent' value='<%=id%>'/>
    </go>
</anchor>
<br/>提示：从&lt;card&gt;到&lt;/wml&gt;
<%end if%>
<br/>----------<br/>
<%if id<>0 then %>
<a href="Clist.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;lxl=<%=lxl%>">[栏目分类]</a><br/>
<%end if%>
<a href="class.asp?sid=<%=sid%>">[栏目管理]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml><%call CloseConn%>