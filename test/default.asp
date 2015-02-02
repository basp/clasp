<!doctype html>	
<html>
<body>
<!--#include file="includes/aspJSON1.17.asp" -->
<!--#include file="includes/http.asp" -->
<%
sub trace(s)
	response.write "<pre>" & s & "</pre>"
end sub

set data = createobject("Scripting.Dictionary")
data.add "foo", formatdatetime(now, vbshortdate)
data.add "bar", 123

dim url: url = "http://localhost/api/json_object.html"
set res = http_get((url), data)
trace(res.data("foo"))
trace(url)
%>
</body>
</html>