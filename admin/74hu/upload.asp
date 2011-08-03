<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<!--#include file="upload.inc.asp"-->
<html>
<head>
<title>WAP2.0文件上传</title>
</head>
<body topmargin="25" leftmargin="20">
<table width=100% border=0 cellspacing="0" cellpadding="0"><tr><td class=tablebody1 width=100% height=100% >
<%
dim upload,file,formName,formPath,filename,fileExt
dim ranNum
call UpFile()
'===========无组件上传(upload_0)====================
sub UpFile()
set upload=new UpFile_Class '建立上传对象
upload.GetData (2097152) '取得上传数据,此处即为2 M

if upload.err > 0 then
select case upload.err
case 1
Response.Write "<script>alert(""请先选择你要上传的文件!"");history.back();</script>"
case 2
Response.Write "<script>alert(""文件大小超过了限制2 M!请使用ftp上传！"");history.back();</script>"
end select
exit sub
else
formPath=upload.form("filepath") '文件保存目录,此目录必须为程序可读写
if formPath="" then
formPath="/upload/"
end if
'在目录后加(/)
if right(formPath,1)<>"/" then 
formPath=formPath&"/"
end if 
for each formName in upload.file '列出所有上传了的文件
set file=upload.file(formName) '生成一个文件对象
if file.filesize<100 then
response.write "<script>alert(""请先选择你要上传的文件！"");history.back();</script>"
response.end
end if

fileExt=lcase(file.FileExt)
if CheckFileExt(fileEXT)=true then
response.write "<script>alert(""文件格式不正确！"");history.back();</script>"
response.end
end if

randomize
ranNum=int(90000*rnd)+10000
filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
if file.FileSize>0 then '如果 FileSize > 0 说明有文件数据
result=file.SaveToFile(Server.mappath(filename)) '保存文件
if result="ok" then
response.write "<br>恭喜，上传成功！<br>文件名：<input name="""&minute(now)&""&second(now)&""" value="""&filename&"""/><br>-------------<br><a href=""fileman.asp?sid="&sid&""">[文件管理]</a><br><a href=""index.asp?sid="&sid&""">[站长工具]</a><br><a href=""../index.asp?sid="&sid&""">[后台管理]</a>"
else
response.write "<br>Sorry，上传失败！"&result&"<br>-------------<br><a href=""fileman.asp?sid="&sid&""">[文件管理]</a><br><a href=""index.asp?sid="&sid&""">[站长工具]</a><br><a href=""../index.asp?sid="&sid&""">[后台管理]</a>"
end if
end if
set file=nothing
next
set upload=nothing
end if
end sub

'判断文件类型是否合格
Private Function CheckFileExt (fileEXT)
dim Forumupload
Forumupload="asp,php,exe,aspx,cgi,html,htm,shtml,jsp"
Forumupload=split(Forumupload,",")
for i=0 to ubound(Forumupload)
if lcase(fileEXT)=lcase(trim(Forumupload(i))) then
CheckFileExt=true
exit Function
else
CheckFileExt=false
end if
next
End Function
%>
</td></tr></table>
</body>
</html>