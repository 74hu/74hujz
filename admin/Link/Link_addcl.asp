<!--#include file="Head.asp"-->
<!--#include file="Config.asp"-->
<!--#include file="mymin.asp"-->
<%
dim classs,pid,br,ID
ID=request.QueryString("ID")
classs=LCase(Request.Form("class"))
pid=LCase(Request.Form("pid"))
br=LCase(Request.Form("br"))
if classs="" or classs=" " then 
Noclass
else
IF Not IsEmpty(Pid) and Not IsNumeric(Pid) Then
Nopid
else
classsave
end if
end if
%><%sub classsave()%><card title="添加友链类别"><p>
<%dim Rs,Sql
call conndata
Set RS=server.createobject("adodb.recordset")
Sql="select * from 74hu_linkc"
RS.open sql,conn,1,3
        RS.addnew
RS("class")=classs
RS("pid")=pid
RS("br")=br
RS.update()
response.write ("友链类别添加成功<br/>")%>
<a href="Link_class.asp?sid=<%=sid%>">[分类管理]</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>[友链管理]</a>
</p></card></wml>
<%Response.End%><%End Sub%>
<%call ALLClose()%>
<%sub Noclass()%><card title="添加友链类别"><p>
<%response.write ("友链类别名称好像不可以为空吧?<br/>")%>
<%response.write ("<a href='Link_add.asp?sid="&sid&"'>返回修改</a><br/>")%>
<a href="Link_class.asp?sid=<%=sid%>">[分类管理]</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>[友链管理]</a>
</p></card></wml>
<%Response.End%><%End Sub%>
<%sub Nopid()%><card title="添加友链类别"><p>
<%response.write ("排序这样写也行呀?<br/>")%>
<%response.write ("<a href='Link_add.asp?sid="&sid&"'>返回修改</a><br/>")%>
<a href="Link_class.asp?sid=<%=sid%>">[分类管理]</a><br/>
<a href='mymin_index.asp?sid=<%=sid%>'>[友链管理]</a>
</p></card></wml>
<%Response.End%><%End Sub%>
