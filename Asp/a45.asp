<HTML>
<BODY>
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

ec=0

response.write "<p>SQL="+sql+"<p>"
rs.open sql, "DSN=gl1197;UID=gl1197;PWD=FDT24tccG;"
rc=0
while not rs.EOF
  rc=rc+1
  major=rs("major")
  minor=rs("minor")
  sub1=rs("sub1")
  sub2=rs("sub2")
  ad=rs("acctdesc")
  bal=rs("balance")
  rs.MoveNext
wend

if rc<>1 then
response.write "Account Not Found!"
ec=ec+1
else
response.write "<table border='2'>"
response.write "<tr>"
response.write "<td>    <font face='COMIC SANS MS,HELVETICA,ARIAL' size='2'>   "
response.write "MAJOR"
response.write "</td>"
response.write "<td> <font face='COMIC SANS MS,HELVETICA,ARIAL' size='2'>  "     
response.write "MINOR"
response.write "</td>"
response.write "<td> <font face='COMIC SANS MS,HELVETICA,ARIAL' size='2'>  "     
response.write "SUB 1"
response.write "</td>"
response.write "<td> <font face='COMIC SANS MS,HELVETICA,ARIAL' size='2'>  "     
response.write "SUB 2"
response.write "</td>"
response.write "<td><font face='COMIC SANS MS,HELVETICA,ARIAL' size='2'>  "  
response.write "ACCOUNT DESCRIPTION"
response.write "</td>"
response.write "<td><font face='COMIC SANS MS,HELVETICA,ARIAL' size='2'>  "  
response.write "BALANCE"
response.write "</td>"

response.write "</tr>"

response.write "<tr>"
response.write "<td>"
response.write cstr(major)+"</td><td>"
response.write cstr(minor)+"</td><td>"
response.write cstr(sub1)+"</td><td>"
response.write cstr(sub2)+"</td><td>"
response.write cstr(ad)+"</td><td>"
response.write cstr(bal)+"</td></tr>"
response.write "</table>"

rs.Close

'response.write ("major="+cstr(major)+"minor="+cstr(minor)+"sub1="+cstr(sub1)+"sub2="+cstr(sub2)+"ad="+cstr(ad)+"bal="+cstr(bal))
end if
%>

<p>
<input type="button" onClick="parent.location='index.htm'" value="Return to Main Menu">
</BODY>
</HTML>