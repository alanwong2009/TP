<HTML>
<BODY>
<p>Alan Wong
<p>0847623
<p>gl1197
<p>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql_string="Select * from je where sourceref="+request.form("sourceref")+" order by sourceref"

response.write "<p>SQL--->"+sql_string+"<---<p>"
rs.open sql_string, "DSN=gl1197;UID=gl1197;PWD=FDT24tccG;"

response.write "<p><table border='1'><tr>"

for i = 0 to rs.fields.count - 1
  response.write "<td align='center'>"+cstr(rs(i).Name)+"</td>"
next
response.write "</tr>"
row_count=0
while not rs.EOF
  row_count=row_count+1
  response.write "<tr>"
  for i = 0 to rs.fields.count - 1
     response.write "<td align='right'>"+cstr(rs(i))+"</td>"
  next
  response.write "</tr>"
  rs.MoveNext
wend

response.write "</table><p>Row Count="+cstr(row_count)
response.write "<p>Normal Termination "+cstr(now)
rs.Close
set rs=nothing
%>
</BODY>
</HTML>