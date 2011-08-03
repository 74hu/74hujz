<!--#include file="co.asp"--><%
'过滤不法字符
function hu(str)
	str=trim(str)
	if IsNull(str) then exit function
	str = replace(str,"'","’")
	str = replace(str,"","")
	str = replace(str," ","")
	str = replace(str,"Λ","")
	str = replace(str,"Ψ","")
	str = replace(str,"","")
	str = replace(str,"%","")
	str = replace(str,"&","")
	str = replace(str,"#","")
	str = replace(str,"*","")
	str = replace(str,"=","")
	str = replace(str,"74hu_","",1,-1,1)
	str = replace(str,"and","",1,-1,1)
	str = replace(str,"%20from","",1,-1,1)
	str = replace(str,"mid","",1,-1,1)
	str = replace(str,"update","",1,-1,1)
	str = replace(str,"exec","",1,-1,1)
	str = replace(str,"select","",1,-1,1)
	str = replace(str,"insert","",1,-1,1)
	str = replace(str,"delete","",1,-1,1)
	str = replace(str,"drop","",1,-1,1)
	str = replace(str,"create","",1,-1,1)
	str = replace(str,"eval","",1,-1,1)
	str = replace(str,"command","",1,-1,1)
	str = replace(str,"dir","",1,-1,1)
	str = replace(str,"truncate","",1,-1,1)
	str = replace(str,"xp_","",1,-1,1)
	str = replace(str,"sp_","",1,-1,1)
	str = replace(str,"master","",1,-1,1)
	str = replace(str,"declare","",1,-1,1)
	str = replace(str,"count","",1,-1,1)
	str = replace(str,"char","",1,-1,1)
	str = replace(str,"unicode","",1,-1,1)
	str = replace(str,"ascii","",1,-1,1)
	str = replace(str,"cmd","",1,-1,1)
	str = replace(str,"法轮","[滤]")
	str = replace(str,"党","[滤]")
	str = replace(str,"奸","[滤]")
	str = replace(str,"穴","[滤]")
	str = replace(str,"龟","[滤]")
	str = replace(str,"淫","[滤]")
	str = replace(str,"裸","[滤]")
	hu=str
end function

'标题和编辑UBB
function ubb(str)
	str=trim(str)
	if IsNull(str) then exit function
	str = replace(str, ">", "&gt;")
	str = replace(str, "<", "&lt;")
	str = replace(str, "＆", "&amp;")
	str = replace(str, "＂", "&quot;")
	str = replace(str, "$", "$$")
	str = replace(str, "??", "？")
	str = replace(str,"","")
	str = replace(str, "", "")
	str = replace(str, "", "")
	str = replace(str,"","")
	str = replace(str," ","")
	str = replace(str,"","")
	str = replace(str,"","")
	str = replace(str,"""","＂")
	str = replace(str, "", "")
	str = replace(str, "'", "’")
	str = replace(str, "`", "")
	str = replace(str, "€", "")
	str = replace(str, "―", "-")
	str = replace(str,"&amp;","←↑→")
	str = replace(str,"&","&amp;")
	str = replace(str,"←↑→","&amp;")
	str = replace(str,"%1A","")
        str = replace(str,"&#xFFE5;","*")
	str = replace(str,"&nbsp;","")
        str = replace(str,"\\","<br/>")
	str = replace(str, chr(01), "")
	str = replace(str, chr(02), "")
	str = replace(str, chr(03), "")
	str = replace(str, chr(04), "")
	str = replace(str, chr(05), "")
	str = replace(str, chr(06), "")
	str = replace(str, chr(07), "")
	str = replace(str, chr(08), "")
	str = replace(str, chr(09), "")
	str = replace(str, chr(10), "<br/>")
	str = replace(str, chr(11), "")
	str = replace(str, chr(12), "")
        str = replace(str, chr(13), "<br/>")
	str = replace(str, chr(14), "")
	str = replace(str, chr(15), "")
	str = replace(str, chr(16), "")
	str = replace(str, chr(17), "")
	str = replace(str, chr(18), "")
	str = replace(str, chr(19), "")
	str = replace(str, chr(20), "")
	str = replace(str, chr(21), "")
	str = replace(str, chr(22), "")
	str = replace(str, chr(23), "")
	str = replace(str, chr(24), "")
	str = replace(str, chr(25), "")
	str = replace(str, chr(26), "")
	str = replace(str, chr(27), "")
	str = replace(str, chr(28), "")
	str = replace(str, chr(29), "")
	str = replace(str, chr(30), "")
	str = replace(str, chr(31), "")
	str = replace(str, chr(34), "&quot;")
	ubb=str
end function

'标题和编辑UBBEDIT
function ubbedit(str)
	str=trim(str)
	if IsNull(str) then exit function
        str = replace(str,"<br/>","\\")
	str = replace(str, ">", "&gt;")
	str = replace(str, "<", "&lt;")
	str = replace(str, "＆", "&amp;")
	str = replace(str, "＂", "&quot;")
	str = replace(str, "$", "$$")
	str = replace(str, "??", "？")
	str = replace(str,"","")
	str = replace(str, "", "")
	str = replace(str, "", "")
	str = replace(str,"","")
	str = replace(str," ","")
	str = replace(str,"","")
	str = replace(str,"","")
	str = replace(str,"""","＂")
	str = replace(str, "", "")
	str = replace(str, "'", "’")
	str = replace(str, "`", "")
	str = replace(str, "€", "")
	str = replace(str, "―", "-")
	str = replace(str,"&amp;","←↑→")
	str = replace(str,"&","&amp;")
	str = replace(str,"←↑→","&amp;")
	str = replace(str,"%1A","")
        str = replace(str,"&#xFFE5;","*")
	str = replace(str,"&nbsp;","")
	str = replace(str, chr(01), "")
	str = replace(str, chr(02), "")
	str = replace(str, chr(03), "")
	str = replace(str, chr(04), "")
	str = replace(str, chr(05), "")
	str = replace(str, chr(06), "")
	str = replace(str, chr(07), "")
	str = replace(str, chr(08), "")
	str = replace(str, chr(09), "")
	str = replace(str, chr(10), "")
	str = replace(str, chr(11), "")
	str = replace(str, chr(12), "")
        str = replace(str, chr(13), "\\")
	str = replace(str, chr(14), "")
	str = replace(str, chr(15), "")
	str = replace(str, chr(16), "")
	str = replace(str, chr(17), "")
	str = replace(str, chr(18), "")
	str = replace(str, chr(19), "")
	str = replace(str, chr(20), "")
	str = replace(str, chr(21), "")
	str = replace(str, chr(22), "")
	str = replace(str, chr(23), "")
	str = replace(str, chr(24), "")
	str = replace(str, chr(25), "")
	str = replace(str, chr(26), "")
	str = replace(str, chr(27), "")
	str = replace(str, chr(28), "")
	str = replace(str, chr(29), "")
	str = replace(str, chr(30), "")
	str = replace(str, chr(31), "")
	str = replace(str, chr(34), "&quot;")
	ubbedit=str
end function

function ubbcode(str)
if IsNull(str) then exit function
str=trim(str)
str = replace(str,"&amp;","←↑→")
str = replace(str,"&","&amp;")
str = replace(str,"←↑→","&amp;")
str = replace(str,"<","&lt;")
str = replace(str,">","&gt;")
str = replace(str,"'","&apos;")
str = replace(str,"""","&quot;")
str = replace(str,"$","$$")
str = replace(str, "", "")
str = replace(str, "", "")
str = replace(str,"","")
str = replace(str," ","")
str = replace(str,"","")
str = replace(str,"","")
str = replace(str,"&#xFFE5;","*")
str = replace(str,"&nbsp;","")
str = replace(str,"[br]","<br/>")
str = replace(str,"[tid]",""&tid&"")
str = replace(str,"[date]",""&date&"")
str = replace(str,"[time]",""&time&"")
str = replace(str,"[now]",""&now()&"")
str = replace(str,"[week]",""&WeekDayName(DatePart("w",Now))&"")
str = replace(str,"[month]",""&Month(Now)&"")
str = replace(str,"[day]",""&Day(Now)&"")
str = replace(str,"[hello]",""&gethello()&"")
str = replace(str,"[favor]",""&getfavor()&"")
str = replace(str,Chr(13),"\\")
str = replace(str,Chr(14),"\\")
str = replace(str, "&amp;quot;", "&quot;")
Set re=new RegExp
re.IgnoreCase =true
re.Global=True
re.pattern="(\[img\])(.[^\[]*)(\[\/img\])"
str = re.Replace(str,"<img src=""$2"" alt='.'/>")
re.pattern="(\[url\])(.[^\[]*)(\[\/url\])"
str = re.Replace(str,"<a href=""$2"" >$2</a>")
re.pattern="(\[url=(.[^\]]*)\])(.[^\[]*)(\[\/url\])"
str = re.Replace(str,"<a href=""$2"" >$3</a>")
re.pattern="(\[c\])(.[^\[]*)(\[\/c\])"
str = re.Replace(str,"<a href=""wtai://wp/mc;$2"" >$2</a>")
re.pattern="(\[c=(.[^\]]*)\))(.[^\[]*)(\[\/c\])"
str = re.Replace(str,"<a href=""wtai://wp/mc;$2"" >$3</a>")
re.pattern="(\[u\])(.[^\[]*)(\[\/u\])"
str = re.Replace(str,"<u>$2</u>")
re.pattern="(\[b\])(.[^\[]*)(\[\/b\])"
str = re.Replace(str,"<b>$2</b>")
re.pattern="(\[i\])(.[^\[]*)(\[\/i\])"
str = re.Replace(str,"<i>$2</i>")
re.Pattern="(\\\\)"
str = re.Replace(str,"<br/>")

set re=Nothing
ubbcode=str
end function

'简单问候语
function getHello()
If Time < #06:00:00# And Time >= #00:30:00# Then 
     getHello="凌晨好！"
ElseIf Time < #09:00:00# And Time >= #06:00:00# Then 
     getHello="早上好！"
ElseIf Time < #11:30:00# And Time >= #09:00:00# Then 
     getHello="上午好！"
ElseIf Time < #12:30:00# And Time >= #11:30:00# Then 
     getHello="中午好！"    
ElseIf Time < #18:00:00# And Time >= #12:30:00# Then
     getHello="下午好！"
ElseIf Time < #20:00:00# And Time >= #18:00:00# Then 
     getHello="傍晚好！"      
ElseIf Time < #23:30:00# And Time >= #20:00:00# Then 
     getHello="晚上好！"    
Else 
     getHello="午夜好！"
End If 
end function

'完整问候语
function getfavor()
If Time < #06:00:00# And Time >= #00:30:00# Then 
     getfavor=""&month(now)&"月"&day(now)&"日"&" "&"凌晨好！"
ElseIf Time < #09:00:00# And Time >= #06:00:00# Then 
     getfavor=""&month(now)&"月"&day(now)&"日"&" "&"早上好！"
ElseIf Time < #11:30:00# And Time >= #09:00:00# Then 
     getfavor=""&month(now)&"月"&day(now)&"日"&" "&"上午好！"
ElseIf Time < #12:30:00# And Time >= #11:30:00# Then 
     getfavor=""&month(now)&"月"&day(now)&"日"&" "&"午饭时间到啦。"    
ElseIf Time < #18:00:00# And Time >= #12:30:00# Then
     getfavor=""&month(now)&"月"&day(now)&"日"&" "&"下午好！"
ElseIf Time < #19:30:00# And Time >= #18:00:00# Then 
     getfavor=""&month(now)&"月"&day(now)&"日"&" "&"晚饭时间到啦。"      
ElseIf Time < #23:30:00# And Time >= #19:30:00# Then 
     getfavor=""&month(now)&"月"&day(now)&"日"&" "&"晚上好！"    
Else
     getfavor=""&month(now)&"-"&day(now)&""&" "&"夜深了,注意休息"
End If
end function

'随机广告
function adstr(adsnum)
dim rsads
set rsads = server.createobject("adodb.recordset")
rsads.open"select  id,name from 74hu_gogo where typeID="&adsnum&" order by id desc ",conn,1,1
if not rsads.eof then
   dim adsranNum
   Randomize()  
adsranNum = int(rsads.recordCount*rnd)+1 
rsads.absoluteposition=adsranNum
Response.Write ("<a href='?aid=url&amp;id="&rsads("id")&"'>"&ubb(rsads("name"))&"</a>")
end if
     rsads.close
set rsads=nothing
end function

'随机广告,定义数目
function adstrs(adsnum,num)
dim rsads
set rsads = server.createobject("adodb.recordset")
Randomize
rsads.open"select top "&num&" id,name from 74hu_gogo where typeID="&adsnum&" order by rnd(-(id+" & rnd() & ")) ",conn,1,1
  while not rsads.EOF
Response.Write ("<a href='?aid=url&amp;id="&rsads("id")&"'>"&ubb(rsads("name"))&"</a><br/>")
rsads.MoveNext
  wend
     rsads.close
set rsads=nothing
end function

'定义广告
Function adsetkf(adnum)
dim rsadset,adset1,adset2,adset3,adset4,adset5
set rsadset=server.CreateObject("adodb.recordset")
rsadset.open"select "&adnum&" from 74hu_control where ID=1",conn,1,1
if not rsadset.eof then
adsetkf=rsadset(adnum)
end if
rsadset.close
set rsadset=nothing
end function

'写入数据库
function usb(str)
	str=trim(str)
	if IsNull(str) then exit function
	str = replace(str,"","")
	str = replace(str," ","")
	str = replace(str,"Λ","")
	str = replace(str,"Ψ","")
	str = replace(str,"","")
	str = replace(str,"'","’")
	str=replace(str,"file:","file：")
	str=replace(str,"files:","files：")
	str=replace(str,"script:","script：")
	str=replace(str,"js:","js：")
        str=replace(str,Chr(10),"\\")
        str=replace(str,Chr(13),"\\")
        str=replace(str,vbnewline,"\\")
        str=replace(str,VbCrLf,"\\")
	usb=str
end function

'最新文章
function newtitle(num,relid)
if relid<>0 then
gettest="where classid="&relid
end if
         set rs1 = server.createobject("adodb.recordset")
rs1.open"select id,title,classid from 74hu_article "&gettest&" order by id desc",conn,1,1
          If rs1.eof Then 
            response.write("还没有文章！<br/>")
                  else
            rs1.Move(0)
                  a=1
            do while ((not rs1.EOF) and a <=num)
               response.write"<a href='?aid=art&amp;id="&rs1("id")&"'>"&ubb(rs1("title"))&"</a><br/>"
                  rs1.MoveNext
                      a=a+1
                            loop
                   end if
                rs1.close
        set rs1=nothing
end function
'最热文章
function hottitle(num,relid)
if relid<>0 then
gettest="where classid="&relid
end if
        set rs2 = server.createobject("adodb.recordset")
         rs2.open"select id,title,classid from 74hu_article "&gettest&" order by hit desc",conn,1,1
          If rs2.eof Then 
            response.write("还没有文章！<br/>")
                  else
            rs2.Move(0)
                 b=1
            do while ((not rs2.eof) and b <=num)
               response.write"<a href='?aid=art&amp;id="&rs2("id")&"'>"&ubb(rs2("title"))&"</a><br/>"
                  rs2.MoveNext
                      b=b+1
                            loop
                   end if
                rs2.close
        set rs2=nothing
end function
'随机文章
function wendtitle(num,relid)
if relid<>0 then
gettest="where classid="&relid
end if
       set rs3 = server.createobject("adodb.recordset")
        Randomize
          rs3.open"select top "&num&" id,title,classid from 74hu_article "&gettest&" order by rnd(-(id*"&rnd()&")) ",conn,1,1
            while not rs3.eof
              response.write"<a href='?aid=art&amp;id="&rs3("id")&"'>"&ubb(rs3("title"))&"</a><br/>"
                  rs3.MoveNext
                      wend
                rs3.close
        set rs3=nothing
end function

'统一时间 2008.8.8 20:08
Function fordate(hu)
if IsDate(hu) = True then
  fordate = year(hu) & "." & month(hu) & "." & day(hu) & " "
  if Hour(hu) < 10 then
  fordate = fordate&"0"
  end if
  fordate = fordate&Hour(hu)&":"
  if Minute(hu) < 10 then
  fordate = fordate&"0"
  end if
  fordate = fordate&Minute(hu)
end if
End Function
'统一时间 8-8 20:08
Function fordate2(hu)
if IsDate(hu) = True then
   fordate2 =""
  fordate2 = fordate2&month(hu) & "-"
  fordate2 = fordate2&day(hu)&" "
  fordate2 = fordate2&Hour(hu)&":"
  if Minute(hu) < 10 then
  fordate2 = fordate2&"0"
  end if
  fordate2 = fordate2&Minute(hu)
end if
End Function
%>