<HTML>
<BODY>
<p>Alan Wong
<p>0847623
<p>gl1197
<p>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql_string="Select * from glmaster order by major asc, minor ASC, sub1 asc,sub2 asc"
sum=0
response.write "<p>SQL--->"+sql_string+"<---<p>"
rs.open sql_string, "DSN=gl1197;UID=gl1197;PWD=FDT24tccG;"

response.write "<p><table border='1'><tr>"

for i = 0 to rs.fields.count - 1
  response.write "<td align='center'>"+cstr(rs(i).Name)+"</td>"
next
response.write "</tr>"
row_count=0
while not rs.EOF
  sum=sum+cstr(rs("balance"))
  row_count=row_count+1
  response.write "<tr>"
  for i = 0 to rs.fields.count - 1
     response.write "<td align='right'>"+cstr(rs(i))+"</td>"
  next
  response.write "</tr>"
  rs.MoveNext
wend

if sum < 0.01 then
 if sum > 0 then
	sum = 0
 else
	sumof=sumof
 end if
end if

if sum > -0.01 then
 if sum < 0 then
	sum = 0
  else
	sumof=sumof
 end if
end if
response.write "<tr><td colspan='5' align='right'>Sum of the Accounts</td><td align='right'>"+cstr(sum)+"</td></tr>"
response.write "</table><p>Row Count="+cstr(row_count)
response.write "<p>Normal Termination "+cstr(now)
rs.Close
set rs=nothing
%>
</BODY>
</HTML>