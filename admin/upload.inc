<%
'----------------------------------------------------------------------
'*******************  无惧上传类 V1.2  *******************************
'作    者:梁无惧
'版权声明:版权所有,源代码公开,各种用途均可免费使用,但是修改后必须把修改后的文件
'发送一份给作者.并且保留作者此版权信息
'**********************************************************************
'----------------------------------------------------------------------
Dim UpFileStream
Class UpFile_Class

	Dim Version,Form,File,Err 

	Private Sub Class_Initialize
		Version="无惧上传组件 V1.2"
		Err = -1
	End Sub

	Private Sub Class_Terminate  
		'清除变量及对像
		If Err < 0 Then
			Form.RemoveAll
			Set Form = Nothing
			File.RemoveAll
			Set File = Nothing
			UpFileStream.Close
			Set UpFileStream = Nothing
		End If
	End Sub
   
	Public Sub GetData(MaxSize)
		'定义变量
		Dim RequestBinDate,sSpace,bCrLf,sInfo,iInfoStart,iInfoEnd,tStream,iStart
		Dim oFileInfo
		Dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
		Dim iFindStart,iFindEnd
		Dim iFormStart,iFormEnd,sFormName
		'代码开始
		If Request.TotalBytes < 1 Then  '如果没有数据上传
			Err = 1
			Exit Sub
		End If
		
		If MaxSize > 0 Then '如果限制大小
			If Request.TotalBytes > MaxSize Then
				Err = 2	'如果上传的数据超出限制
				Exit Sub
			End If
		End If
		
		Set Form = Server.CreateObject ("Scripting.Dictionary")
		Form.CompareMode = 1
		Set File = Server.CreateObject ("Scripting.Dictionary")
		File.CompareMode = 1
		Set tStream = Server.CreateObject ("ADODB.Stream")
		Set UpFileStream = Server.CreateObject ("ADODB.Stream")
		UpFileStream.Type = 1
		UpFileStream.Mode = 3
		UpFileStream.Open 
		UpFileStream.Write Request.BinaryRead (Request.TotalBytes)
		UpFileStream.Position = 0
		RequestBinDate = UpFileStream.Read 
		iFormEnd = UpFileStream.Size
		bCrLf = ChrB (13) & ChrB (10)
		
		'取得每个项目之间的分隔符
		sSpace = MidB (RequestBinDate,1, InStrB (1,RequestBinDate,bCrLf)-1)
		iStart = LenB  (sSpace)
		iFormStart = iStart+2
		
		'分解项目
		Do
			iInfoEnd = InStrB (iFormStart,RequestBinDate,bCrLf & bCrLf)+3
			tStream.Type = 1
			tStream.Mode = 3
			tStream.Open
			UpFileStream.Position = iFormStart
			UpFileStream.CopyTo tStream,iInfoEnd-iFormStart
			tStream.Position = 0
			tStream.Type = 2
			tStream.CharSet = "UTF-8"
			sInfo = tStream.ReadText
				 
			'取得表单项目名称
			iFormStart = InStrB (iInfoEnd,RequestBinDate,sSpace)-1
			iFindStart = InStr (22,sInfo,"name=""",1)+6
			iFindEnd = InStr (iFindStart,sInfo,"""",1)
			sFormName = Mid  (sinfo,iFindStart,iFindEnd-iFindStart)
			
			'如果是文件
			If InStr(45,sInfo,"filename=""",1) > 0 Then
				Set oFileInfo = New FileInfo_Class
				'取得文件属性
				iFindStart = InStr (iFindEnd,sInfo,"filename=""",1)+10
				iFindEnd = InStr (iFindStart,sInfo,"""",1)
				sFileName = Mid  (sinfo,iFindStart,iFindEnd-iFindStart)
				oFileInfo.FileName = Mid (sFileName,InStrRev (sFileName, "\")+1)
				oFileInfo.FilePath = Left (sFileName,InStrRev (sFileName, "\"))
				oFileInfo.FileExt = Mid (sFileName,InStrRev (sFileName, "."))
				iFindStart = InStr (iFindEnd,sInfo,"Content-Type: ",1)+14
				iFindEnd = InStr (iFindStart,sInfo,vbCr)
				oFileInfo.FileType = Mid  (sinfo,iFindStart,iFindEnd-iFindStart)
				oFileInfo.FileStart = iInfoEnd
				oFileInfo.FileSize = iFormStart -iInfoEnd -2
				oFileInfo.FormName = sFormName
				file.add sFormName,oFileInfo
			Else
				'如果是表单项目
				tStream.Close
				tStream.Type = 1
				tStream.Mode = 3
				tStream.Open
				UpFileStream.Position = iInfoEnd 
				UpFileStream.CopyTo tStream,iFormStart-iInfoEnd-2
				tStream.Position = 0
				tStream.Type = 2
				tStream.CharSet = "UTF-8"
				sFormValue = tStream.ReadText
				
				If Form.Exists (sFormName) Then
					Form (sFormName) = Form (sFormName) & ", " & sFormValue
				Else
					form.Add sFormName,sFormValue
				End If
				
			End If
			tStream.Close
			iFormStart = iFormStart+iStart+2
			'如果到文件尾了就退出
		Loop Until (iFormStart+2) >= iFormEnd 
		RequestBinDate = ""
		Set tStream = Nothing
	End Sub
End Class
'----------------------------------------------------------------------------------------------------
'文件属性类
Class FileInfo_Class
	Dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
	'保存文件方法
	Public Function SaveToFile(Path)
		On Error Resume Next
		Dim oFileStream
		Set oFileStream = CreateObject ("ADODB.Stream")
		oFileStream.Type = 1
		oFileStream.Mode = 3
		oFileStream.Open
		UpFileStream.Position = FileStart
		UpFileStream.CopyTo oFileStream,FileSize
		oFileStream.SaveToFile Path,2
		oFileStream.Close
		Set oFileStream = Nothing 
	End Function
 
	'取得文件数据
	Public Function FileData
		UpFileStream.Position = FileStart
		FileData = UpFileStream.Read (FileSize)
	End Function

End Class
%>