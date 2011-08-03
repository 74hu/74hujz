<%
Function CodeWML(STR)
	IF IsNull(STR) THEN EXIT Function
	STR=REPLACE(STR,"&","&amp;")
	STR=REPLACE(STR,"<","&lt;")
	STR=REPLACE(STR,">","&gt;")
	STR=REPLACE(STR,"$","$$")
	STR=REPLACE(STR,"'","&apos;")
	STR=REPLACE(STR,"""","&quot;")
	STR=REPLACE(STR,"&nbsp;"," ")
	STR=REPLACE(STR,"&amp;lt;","&lt;")
	STR=REPLACE(STR,"&amp;gt;","&gt;")
	STR=REPLACE(STR,"&amp;amp;","&amp;")
	CodeWML=STR
End Function


Function XMLHTTPGet(URL) 
	Dim XMLGet
	Set XMLGet=Server.CreateObject("Microsoft.XMLHTTP") 
	XMLGet.Open "GET",URL,False
	XMLGet.SetRequestHeader "USER-AGENT","Shine System V 1.0"
	XMLGet.SetRequestHeader "ACCEPT","*/*,http://74hu.cn"
	XMLGet.Send() 

	XMLHTTPGet=UTF8Tran(XMLGet.ResponseBody,"UTF-8")
	Set XMLGet=nothing

End Function

Function UTF8Tran(ServerInfo,Cset) 
	Dim TranStream
	Set TranStream = Server.CreateObject("Adodb.Stream") 
	TranStream.Type = 1
	TranStream.Mode =3
	TranStream.Open
	TranStream.Write ServerInfo
	TranStream.Position = 0
	TranStream.Type = 2
	TranStream.Charset = Cset
	UTF8Tran = TranStream.ReadText
	TranStream.Close
	Set TranStream = nothing
End Function

Function GetText(STR)
	Dim TextRE,TextContents,Contents,NewContent
	Set TextRE = New Regexp
	TextRE.IgnoreCase = False
	TextRE.Global = True
	TextRE.Pattern = "<[^>]+>?[^<]*>"
	Set TextContents=TextRE.Execute(STR)
		For Each Contents In TextContents
			STR=Trim(Replace(STR,Contents.Value,""))
			'STR=Contents.Value
		Next
	GetText=STR
End Function
Function RemoveHTML(strHTML) 
Dim objRegExp, Match, Matches 
Set objRegExp = New Regexp 

objRegExp.IgnoreCase = True 
objRegExp.Global = True 
'取闭合的<> 
objRegExp.Pattern = "<.+?>" 
'进行匹配 
Set Matches = objRegExp.Execute(strHTML) 

' 遍历匹配集合，并替换掉匹配的项目 
For Each Match in Matches 
strHtml=Replace(strHTML,Match.Value,"") 
Next 
RemoveHTML=strHTML 
Set objRegExp = Nothing 
End Function
%>