<!-- #include file="ding.asp" -->
<!-- #include file="mymin.asp" -->
<%Call Head()%>
<card title="处理WML页面"><p>
<%  
IF KEY<>0 then
  Call Error("你的权限不足！")
  end if
dim id,lxl,class1,url,Wmlname,wmltxt
id=request("id")
if id="" or IsNumeric(id)=False then ID=0
url=request.form("url")
Wmlname=request.form("Wmlname")
wmltxt=request.form("wmltxt")
if wmltxt="" then
  Call Error("WML页面内容不能为空！")
  end if

active=request.form("active")
if active="" or IsNumeric(active)=False then active=0
        lxl= cint(request("lxl"))
        class1=request("class1")
        relid=request("relid")
        num=request("num")
if class1="" then
        Call Error("栏目名称不能为空！")
        end if
if relid="" or IsNumeric(relid)=False then
 num=0
  end if
if num="" or IsNumeric(num)=False then num=0
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
  end if

	dim read,reads,fso,filesize
        if url<>"" then
                filename=url
        else
        if Wmlname="" then
        filename="/wml/"&addwml(now())
        else
        filename="/wml/"&Wmlname&".wml"
        end if
        end if
	read = "wml.txt"
	reads=LoadFile("tools/wml.txt")
	call SaveToFile(reads&wmltxt,filename)
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filesize=fso.GetFile(Server.MapPath(filename)).size
call conndata
	set rs=server.createobject("adodb.recordset")
       if url<>"" then
	sql="select * from class where classid="&id 
	else
	sql="select * from class" 
	end if
	rs.open sql,conn,2,3
        if url<>"" then
        else
	rs.addnew
        end if
	rs("class")=class1
        rs("pid")=pid
        rs("br")=br
        rs("lx")=lx
        if parent<>0 then
        rs("parent")=parent
        end if
        rs("url")=filename
        rs("active")=active
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
if url<>"" then
response.write"WML页面编辑成功！"
else
response.write"WML页面添加成功！"
end if%>
<br/>预览WML<a href='<%=filename%>'><%=filename%></a><br/>
<%if id<>0 then %>
<a href="Clist.asp?sid=<%=sid%>&amp;id=<%=id%>&amp;lx=20">[栏目分类]</a><br/>
<%end if%>
<a href='class.asp?sid=<%=sid%>'>[栏目分类]</a><br/>
<a href="index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card>
</wml><%call CloseConn%>