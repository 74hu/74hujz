<%
Sub Head() 
    Response.ContentType = "text/vnd.wap.wml"
    Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>"
    Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">"
    Response.Write "<wml>"
    Response.Write "<head>"
    Response.Write "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>"
    Response.Write "<meta http-equiv=""Cache-Control"" content=""no-cache""/>"
    Response.Write "</head>"
End Sub

Sub Last()
response.write "<br/><a href=""../index.asp?sid="&sid&""">后台管理</a>"
response.write "<br/><a href=""/index.asp"">网站首页</a><br/>"
response.write "</p></card></wml>"
    Response.End
End Sub

Sub rootLast()
response.write "<br/><a href=""index.asp?sid="&sid&""">后台管理</a>"
response.write "<br/><a href=""/index.asp"">网站首页</a><br/>"
response.write "</p></card></wml>"
    Response.End
End Sub

Function adsetkf(adnum)
		Set conn = Server.CreateObject("ADODB.Connection")
	connstr="driver={Microsoft Access Driver (*.mdb)};pwd=dbq=" & Server.MapPath(""&db&"")
	conn.Open connstr
dim rsadset,adset1,adset2,adset3,adset4,adset5
set rsadset=server.CreateObject("adodb.recordset")
rsadset.open"select "&adnum&" from 74hu_control where ID=1",conn,1,1
if not rsadset.eof then
adsetkf=rsadset(adnum)
end if
rsadset.close
set rsadset=nothing
end function

Function conndata()
		Set conn = Server.CreateObject("ADODB.Connection")
	connstr="driver={Microsoft Access Driver (*.mdb)};pwd=;dbq=" & Server.MapPath(""&db&"")
	conn.Open connstr
end function

Function Error(erstr)
  conn.close
  set conn=Nothing
  Response.write erstr & chr(13)
  Response.write "<br/><anchor><prev/>&#x8FD4;&#x56DE;</anchor>" & chr(13)
  Response.write "</p></card></wml>"
  Response.end
end Function

Function CreatedTextFiles(FileName,body)
On Error Resume Next
If InStr(FileName, ":") = 0 Then FileName = Server.MapPath(FileName)
	Dim oStream
	Set oStream = CreateObject("ADODB.Stream")
	oStream.Type = 2 '设置为可读可写
	oStream.Mode = 3 '设置内容为文本
	oStream.Charset = "utf-8"
	oStream.Open
	oStream.Position = oStream.Size
	oStream.WriteText body
	oStream.SaveToFile FileName, 2
	oStream.Close
	Set oStream = Nothing
	If Err.Number <> 0 Then Err.Clear
End Function
%>