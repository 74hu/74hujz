<%
'================================================
'函数名：addwml
'作  用：时间命名的函数
'参  数：fname -文件路径, str_内容
'================================================

function addwml(fname)
fname = fname '前fname为变量，后fname为函数参数引用
fname = replace(fname,"-","")
fname = replace(fname," ","") 
fname = replace(fname,":","")
fname = replace(fname,"PM","")
fname = replace(fname,"AM","")
fname = replace(fname,"上午","")
fname = replace(fname,"下午","")
addwml = fname & ".wml"
end function 
'================================================
'================================================
'作  用:生成WML文件
'================================================
	Function LoadFile(File)		'文件内容读取.
	Dim objStream
	On Error Resume Next
	Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
		Response.Write "非常遗憾,您的主机不支持ADODB.Stream,不能使用本程序"
		Err.Clear
'		Response.End
		End If
	With objStream
	.Type = 2
	.Mode = 3
	.Open
	.LoadFromFile Server.MapPath(File)
		If Err.Number<>0 Then
		Response.Write "文件"&File&"无法被打开，请检查是否存在!"
		Err.Clear
'		Response.End
		End If
	.Charset = "utf-8"
	.Position = 2
	LoadFile = .ReadText
	.Close
	End With
	Set objStream = Nothing
	End Function

	Sub SaveToFile(strBody,File)		'存储内容到文件
	Dim objStream
	On Error Resume Next
	Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number=-2147221005 Then 
		Response.Write "非常遗憾,您的主机不支持ADODB.Stream,不能使用本程序"
		Err.Clear
'		Response.End
		End If
	With objStream
	.Type = 2
	.Open
	.Charset = "utf-8"
	.Position = objStream.Size
	.WriteText = strBody
	.SaveToFile Server.MapPath(File),2
	.Close
	End With
	Set objStream = Nothing
	End Sub
%>