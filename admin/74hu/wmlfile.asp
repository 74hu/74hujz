<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>WAP2.0上传WML文件</title>
</head>
<body>
<%
conn.close
set conn=nothing
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-cache, must-revalidate"
%>
<%
response.write("<form action='Wml_Upfile.asp?sid="&sid&"' enctype='multipart/form-data' method='post'>")
%>
WML地址命名:(留空则自动命名)<br/><input type="text" name="content" maxlength="30"><br/> 
选择WML文件:<br/><input type="file" name="filefield1"><br/> 
<input type="submit" name="Submit2" value="确定上传"></form>
<br/>1.建议所上传文件名和目录组合是字母与数字.
<br/><a href='wmltext.asp?sid=<%=sid%>'>[WML页面管理]</a>
<br/><a href="index.asp?sid=<%=sid%>">[站长工具]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</body></html>