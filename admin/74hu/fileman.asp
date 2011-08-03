<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<%Call Head()%>
<card title="文件管理"><p>
<%
IF KEY<>0 then
  response.write"你的权限不足！</p></card></wml>"
  response.end
  end if
UploadDir="/upload/"
Const MaxPerPage=10     
'检查组件是否已经安装
Function IsObjInstalled(strClassString)	
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
dim strFileName
dim totalPut,CurrentPage,TotalPages
dim TruePath,fso,theFolder,theFile,whichfile,thisfile,FileCount,TotleSize
strFileName="?"

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

TruePath=Server.MapPath(UploadDir)
If not IsObjInstalled("Scripting.FileSystemObject") Then
	Response.Write "你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能"
Else
	set fso=CreateObject("Scripting.FileSystemObject")

  if fso.FolderExists(TruePath)then
	FileCount=0
	TotleSize=0
	Set theFolder=fso.GetFolder(TruePath)
	For Each theFile In theFolder.Files
		FileCount=FileCount+1
		TotleSize=TotleSize+theFile.Size
	next
    totalPut=FileCount
	if currentpage<1 then
   		currentpage=1
   	end if
   	if (currentpage-1)*MaxPerPage>totalput then
		if (totalPut mod MaxPerPage)=0 then
	  		currentpage= totalPut \ MaxPerPage
	  	else
	      	currentpage= totalPut \ MaxPerPage + 1
		end if

    end if
	if currentPage=1 then
		showContent     	
		showpage2 strFileName,totalput,MaxPerPage
		response.write "<br/>本页共显示" & FileCount-1 & "个文件，占用" & TotleSize\1024 & " K"
   	else
   	   	if (currentPage-1)*MaxPerPage<totalPut then
			showContent     	
			showpage2 strFileName,totalput,MaxPerPage
			response.write "<br/>本页共显示" & FileCount-1 & "个文件，占用" & TotleSize\1024 & "K"
       	else
        	currentPage=1
			showContent     	
			showpage2 strFileName,totalput,MaxPerPage
			response.write "本页共显示" & FileCount-1 & "个文件，占用" & TotleSize\1024 & "K"
    	end if
	end if
  else
	response.write "找不到文件夹！可能是配置有误！"
  end if
end if

sub showContent()
   	dim c
	FileCount=1
	TotleSize=0
%>
      
        <% For Each theFile In theFolder.Files
	c=c+1
	if FileCount>MaxPerPage then
		exit for
	elseif c>MaxPerPage*(CurrentPage-1) then %>
[<a href="filedel.asp?path=<%=(UploadDir & theFile.Name)%>&amp;pathname=<%=theFile.Name%>&amp;TP=<%=TP%>&amp;sid=<%=sid%>">删</a>]<%=C+(CurrentPage-1)*MaxPerPage%>.<a href="<%=(UploadDir &theFile.Name)%>"><%=(UploadDir & theFile.Name)%></a><br/>
<% if FileCount mod 5 =0 then%>

                <%end if%>
        <%FileCount=FileCount+1
	TotleSize=TotleSize+theFile.Size
	end if
Next
%>		
      <%
end sub
%>
    
<%
sub showpage2(sfilename,totalnumber,maxperpage)
	dim n, i,strTemp
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  
  	if n-currentpage<1 then
  	else
    		strTemp=strTemp & "<a href='fileman.asp?page=" & (CurrentPage+1) & "&amp;TP="&TP&"&amp;sid="&sid&"'>下一页</a>&nbsp;"
  	end if
  	if CurrentPage<2 then
  	else
    		strTemp=strTemp & "<a href='fileman.asp?page=" & (CurrentPage-1) & "&amp;TP="&TP&"&amp;sid="&sid&"'>上一页</a>&nbsp;"
  	end if

   	strTemp=strTemp & "<br/>(" & CurrentPage & "/" & n & ") "
	strTemp=strTemp & "共" & totalnumber & "个"

  	if n>1 then
        strTemp=strTemp & "<input name=""page"" format=""*N"" value=""2"" type=""text"" maxlength=""5"" emptyok=""true"" size=""2""/><a href=""fileman.asp?page=$(page)&amp;TP="&TP&"&amp;sid="&sid&""">页</a>"
  	end if
    response.write strTemp
end sub
%>
<br/><a href="Files.asp?sid=<%=sid%>">[文件管理]</a>
<br/><a href="index.asp?sid=<%=sid%>">[站长工具]</a>
<br/><a href='../index.asp?sid=<%=sid%>'>[后台管理]</a></p>
</card>
</wml>