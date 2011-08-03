<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="添加栏目类别">
<p>
<%  IF KEY<>0 then
  Call Error("你的权限不足！")
  end if
 id= cint(request("id"))
        lxl= cint(request("lxl"))
        class1=request("class1")
        relid=request("relid")
        num=request("num")
if class1="" then
        Call Error("栏目名称不能为空！")
        end if
if relid="" or IsNumeric(relid)=False then
 relid=0
  end if
if num="" or IsNumeric(num)=False then
 num=0
  end if
        pid=request.form("pid")
if pid="" or IsNumeric(pid)=False then
  Call Error("排序无效！")
  end if
        br=request.form("br")
if br="" or IsNumeric(br)=False then
  Call Error("换行无效！")
  end if
  lx=request.form("lx")
if lx="" or IsNumeric(lx)=False then
  Call Error("类型无效！")
  end if
        parent=request.form("parent")
        if parent="" or IsNumeric(parent)=False then parent=0
        lid=request.form("lid")
 if lid <>"" then
if IsNumeric(lid)=False then
  Call Error("ID无效！")
  end if
  relid=lid
  end if
        wmltxt=request.form("wmltxt")
	rs.close
	set rs=nothing
	call CloseConn
call conndata
set rs=server.createobject("adodb.recordset")
rs.open "select * from 74hu_class",conn,2,3
	rs.addnew
	rs("class")=class1
        rs("pid")=pid
        rs("br")=br
        rs("lx")=lx
        rs("parent")=parent
        rs("relid")=relid
        rs("wmltxt")=wmltxt
        rs("num")=num
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
栏目添加成功!
<br/>----------<br/>
<%if id<>0 then %>
<a href="Clist.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;lx=<%=lxl%>">[栏目分类]</a><br/>
<%end if%>
<a href='class.asp?sid=<%=sid%>'>[栏目分类]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p>
</card>
</wml>