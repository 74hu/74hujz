<!-- #include file="h.asp" -->
<card title='发表评论'>
<%
o=request.QueryString("o")
if o="" or IsNumeric(o)=False then
o=0
end if

if o=0 then
pl=hu(request.QueryString("pl"))

p=request.QueryString("p")
if p="" or IsNumeric(p)=False then
p=1
end if

if pl="" then
response.write("<p>评论内容不能为空！</p></card></wml>")
response.end
  end if
if len(pl)>100 then
response.write("<p>评论内容最多100字！</p></card></wml>")
response.end
  end if

set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select * from 74hu_pl",conn,1,3
rs.addnew
rs("pl")=pl
rs("ip")=User_Ip
rs("smsid")=id
rs.update

'更新评论
dim Counts
Counts=conn.execute("Select count(ID) from 74hu_pl where smsid="&id&"")(0)
conn.execute("update 74hu_article Set smspin = "&Counts&" where ID="&id)
else
ids=request.QueryString("ids")
if ids="" or isnumeric(ids)=false or isnull(ids) then
ids=1
end if

set rs=Server.CreateObject("ADODB.Recordset")
rs.open"select ag,da from 74hu_pl where id="&ids,conn,1,3
if o=1 then
rs("ag")=rs("ag")+1
else
rs("da")=rs("da")+1
end if
rs.update
end if
response.write("<onevent type='onenterforward'><go href='?aid=dis&amp;id="&id&"&amp;p="&p&"'/></onevent><p>评论发表成功！<br/>")

rs.close
Set rs=nothing

%><a href='?aid=dis&amp;id=<%=id%>&amp;p=<%=p%>'>查看评论</a><br/>
<a href='?aid=art&amp;id=<%=id%>&amp;p=<%=p%>'>返回原文</a><br/>