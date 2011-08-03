<%end if
conn.close
set conn=nothing
response.write "<br/><a href='?aid=guest'>&gt;&gt;给我们提意见</a><br/><a href='?aid=index'>"&waptitle&"</a>-<a href='?aid=map'>导航</a>-<a href='?aid=shuqian'>收藏</a><br/>"
response.write""&wapbei&"</p></card></wml>"
response.end%>