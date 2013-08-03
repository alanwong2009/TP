<html>
<body>
<p>Alan Wong
<p>0847623
<p>gl1197
<p>
<%
set rs=Server.CreateObject("ADODB.Recordset")

sql="Select * from glmaster where major="
sql=sql+request.form("major")+" and minor="
sql=sql+request.form("minor")+" and sub1="
sql=sql+request.form("sub1")+" and sub2="
sql=sql+request.form("sub2")

response.write "<p>SQL--->"+sql+"<---<p>"
rs.open sql, "DSN=gl1197;UID=gl1197;PWD=FDT24tccG;"
rc=0
while not rs.EOF
  rc=rc+1
  rs.MoveNext
wend

if rc<>0 then
	response.write "Account Already exist!"
else

set cn=server.createobject("ADODB.Connection")

response.write "<p>Created Connection named cn OK<p>"

SQL="INSERT INTO GLMASTER VALUES("
SQL=SQL+REQUEST.FORM("major")+" ,"
SQL=SQL+request.form("minor")+" ,"
SQL=SQL+request.form("sub1")+" ,"
SQL=SQL+request.form("sub2")+" ,"
SQL=SQL+chr(39)+request.form("acctdesc")+chr(39)+" , 0.00)"

response.write "<P>SQL="+SQL+"<p>"

cn.open "gl1197","gl1197","FDT24tccG"

response.write "<p>Opened OK<p>"

cn.execute SQL,numa

if numa=1 then

response.write "<P>INSERT OK"

else

response.write "Insert Failed"
end if
end if
rs.Close

%>
<p>
<input type="button" onClick="parent.location='index.htm'" value="Return to Main Menu">

</body>
</html>