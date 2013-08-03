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

response.write "<p>SQL--->"+sql+"<---<p>"
rs.open sql, "DSN=gl1197;UID=gl1197;PWD=FDT24tccG;"
rc=0
while not rs.EOF
  rc=rc+1
  bal=rs("balance")
  rs.MoveNext
wend

if rc<>1 then
response.write "Account Not Found!"
ec=ec+1
else
response.write "Account found, balance = "+cstr(bal)+"<br>"
	if cstr(bal)<>0 then
		ec=ec+1
		response.write ("You cannot delete non-zero balance account")
	else
		if ec=0 then
		response.write "Proceed to delete the account...<br>"
		set cn=server.createobject("ADODB.connection")
		cn.open "gl1197","gl1197","FDT24tccG"
		sql="delete from glmaster where major="
		sql=sql+request.form("major")+" and minor="
		sql=sql+request.form("minor")+" and sub1="
		sql=sql+request.form("sub1")+" and sub2="
		sql=sql+request.form("sub2")

		cn.execute sql,numa
			if numa = 1 then
				response.write ("Delete OK!<br>")
			else
				response.write("Delete Fail!<br>")
			end if
		cn.close
		end if
	end if
end if
rs.Close

%>
<p>
<input type="button" onClick="parent.location='index.htm'" value="Return to Main Menu">
</BODY>
</HTML>