<!--#include file="Head.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="mymin.asp"-->
<card title="添加友链类别">
<p>类别名称:<input name="class<%=minute(now)%><%=second(now)%>" title="名称" maxlength="10" size="10" value=""/><br/>
类别排序:<input name="pid<%=minute(now)%><%=second(now)%>" type="text"   value="" size="10" maxlength="255"/><br/>
		类别换行:<select name="br<%=minute(now)%><%=second(now)%>" value="1">
			<option value="1">自动换行</option>
			<option value="2">不换行</option>
			</select><br/>
<anchor>确认提交
<go href="Link_addcl.asp?sid=<%=sid%>" method="post" accept-charset="utf-8">
<postfield name="class" value="$(class<%=minute(now)%><%=second(now)%>)"/>
<postfield name="pid" value="$(pid<%=minute(now)%><%=second(now)%>)"/>
<postfield name="br" value="$(br<%=minute(now)%><%=second(now)%>)"/>
</go></anchor><br/>----------<br/>
<a href="Link_class.asp?sid=<%=sid%>">[分类管理]</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>[友链管理]</a>
</p></card></wml>
