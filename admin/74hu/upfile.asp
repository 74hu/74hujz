<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>WAP2.0文件上传</title>
</head> 
<body topmargin="25" leftmargin="20">
<%IF KEY<>0 then
  response.write"你的权限不足！请联系管理员</body></html>"
  response.end
  end if%>
请选择要上传的文件 
<form action="upload.asp?sid=<%=sid%>" method="post" enctype="multipart/form-data" name="form1">
<input type="file" name="file">
<!--<br>
<input type="file" name="file">
<br>
<input type="file" name="file">-->
<br><br>
<input type="submit" name="Submit" value="提交"> <input type="reset" name="Submit" value="重置" />
</form>
建议上传200k以下的文件，有些服务器有上传限制！超过2M上传服务器容易超时，请选择2M以下的文件<br>
-------------<br>
<a href="files.asp?sid=<%=sid%>">[文件管理]</a><br>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</body>
</html>