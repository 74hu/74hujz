<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="分类修改结果">
<p>
<%   
IF KEY<>0 then
  Call Error("你的权限不足！")
  end if
        num=request.form("num")
        relid=request.form("relid")

if num="" or IsNumeric(num)=False then
 num=0
  end if
  
if relid="" or IsNumeric(relid)=False then
 num=0
  end if

        id=request.querystring("id")
if id="" or IsNumeric(id)=False then
        Call Error("ID无效！")
        end if

        class1=request.form("class1")
if class1="" then
        Call Error("栏目名称不能为空！")
        end if

        lx=request.form("lx")
if lx="" or IsNumeric(lx)=False then
  Call Error("类型无效！")
  end if

        pid=request.form("pid")
if pid="" or IsNumeric(pid)=False then
  Call Error("排序无效！")
  end if

        br=request.form("br")
if br="" or IsNumeric(br)=False then
  Call Error("换行无效！")
  end if
        wmltxt=request.form("wmltxt")
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_class where classid="&id,conn,1,3
        rs("class")=class1
        rs("pid")=pid
        rs("br")=br
        rs("lx")=lx
        rs("relid")=relid
        rs("wmltxt")=wmltxt
        rs("num")=num
        rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "修改成功!"
%>
<br/>----------<br/>
<%if id<>0 then %>
<a href="Clist.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;lx=<%=lx%>">[栏目分类]</a><br/>
<%end if%>
<a href='class.asp?sid=<%=sid%>'>[栏目分类]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>