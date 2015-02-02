<!--#include file="includes/aspJSON1.17.asp" -->
<!--#include file="includes/http.asp" -->
<!doctype html>	
<html>
<body>
<%
sub trace(s)
	response.write "<pre>" & s & "</pre>"
end sub

set data = createobject("Scripting.Dictionary")
data.add "foo", formatdatetime(now, vbshortdate)
data.add "bar", 123

dim url: url = "http://localhost/api/json_object.html"

'Invoke with parenthesis and `url` is passed by value
set res = http_get((url), data)
trace(url)

'Invoke without parenthesis and `url` is passed by refeference
set res = http_get(url, data)
trace(url)
%>
</body>
</html>