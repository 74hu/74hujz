<!-- #include file="../ding.asp" -->
<!-- #include file="../mymin.asp" -->
<!--#include file="upload.inc"-->
<%Call Head()%>
<card id="main" title="WAP2.0上传WML文件"><p>
<%
	' 忽略所有错误
    on error resume next
Server.ScriptTimeOut = 1800
dim constPath,mypath
mypath=server.mappath("Wml_Upfile.asp")
constPath=replace(mypath,"Wml_Upfile.asp","")
dim upload,oFile,formName,SavePath,filename,fileExt,oFileSize,sizes
dim EnableUpload
dim arrUpFileType
dim ranNum
dim msg,FoundErr,MaxFileSize
MaxFileSize="5120"'最大上传文件，以KB为单位
msg=""
FoundErr=false
EnableUpload=true
dim strMonth,strDay
dim tid,action,bbsid,title,content,content1,content2,content3,content4,content5
kk=0
	set upload=new upfile_class ''建立上传对象
	upload.GetData(5357600)   '取得上传数据,限制最大上传100M
	if upload.err > 0 then  '如果出错
		select case upload.err
			case 1
				msg= "请先选择你要上传的文件！"
			case 2
				msg= "你上传的文件总大小超出了最大限制（5M）"
		end select
		'showWML()


		response.end
	end if

  SavePath = "/wml/"
  	content=forbbs(upload.form("content"))
if len(content1)>30 or len(content2)>30 or len(content3)>30 or len(content4)>30 or len(content5)>30 then
  Call Error("命名最多30字，请返回重试！")
  end if

		
for each formName in upload.file 
		set ofile=upload.file(formName)  '生成一个文件对象	
		upfilename=ofile.FileName		
		oFileSize=ofile.filesize	
		sizes=cstr(round(oFileSize/1024))		
		fileExt=lcase(ofile.FileExt)
   
		if fileEXT<>".wml" then
		 EnableUpload=false
		end if
		if EnableUpload=false then
			msg="这种文件类型不允许上传:asp|asa|aspx|exe|bat|..."
			FoundErr=true
           call error("只支持WML格式文件！")
		  response.end
		end if
		if oFileSize>(MaxFileSize*1024) then
      msg="文件大小超过了限制，最大只能上传" & CStr(MaxFileSize) & "K的文件！"
			FoundErr=true
			'showWML()
	    msg=msg & "<br/><a href=""wmltext.asp?sid="&sid&""">返回上级</a><br/>"
             response.write msg
             response.write "</p></card></wml>"

		  response.end
		end if
		
		
		if FoundErr<>true then
			randomize
			ranNum=int(900*rnd)+100
			
			strMonth=month(now)
			if len(strMonth)=1 then
				strMonth="0"&strMonth
			end if
			strDay=day(now)
			if len(strDay)=1 then
				strDay="0"&strDay
			end if
			
			'--------------------
			
			dim wmlname
	
		    if Content<>"" Then
		    wmlname=Content
		    else
			wmlname=year(now)&strMonth&strDay&hour(now)&minute(now)&second(now)&ranNum
			end if
			
			
			filename=SavePath&wmlname&""&fileExt

      dim realpath
      realpath=filename

ofile.SaveToFile Server.mappath(realpath)   '保存文件   
			kk=kk+1


END IF
    next
	set upload=nothing

		
	    msg="上传WML文件成功！<br/><a href=""wmltext.asp?sid="&sid&""">WML页面管理</a><br/>"

		 conn.close
set conn=nothing
			
%>
<%=msg%>
<a href="index.asp?sid=<%=sid%>">[站长工具]</a><br/>
<a href="../index.asp?sid=<%=sid%>">[后台管理]</a>
</p></card></wml>