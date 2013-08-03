<html>
<body>
<p>Alan Wong
<p>0847623
<p>gl1197
<p>
<%
response.write "Begin...<p>"
nr=0
	for i=0 to 6
		os=request.form("a00"+cstr(i)+"1field")
		if os <> "" then
			nr=nr+1
		end if
	next

response.write "Number of entries="+cstr(nr)+"<p>"

ec=0

set rs=server.createobject("ADODB.recordset")
	for j=0 to nr-1
		sql="select * from glmaster where "
		sql=sql+"major="+request.form("a00"+cstr(j)+"1field")+" and "
		sql=sql+"minor="+request.form("a00"+cstr(j)+"2field")+" and "
		sql=sql+"sub1="+request.form("a00"+cstr(j)+"3field")+" and "
		sql=sql+"sub2="+request.form("a00"+cstr(j)+"4field")
		response.write sql+"<br>"
	

rs.open sql,"DSN=gl1197;UID=gl1197;PWD=FDT24tccG;"

response.write "<p>Open sql OK<p>"

rc=0
	while not rs.eof
		rc=rc+1
		rs.movenext
	wend

rs.close


	if rc<>1 then
		ec=ec+1
		response.write "Account Not Found!<p>"
	end if
next

sqlsrn="select * from je where sourceref="+request.form("sourceref")

rs.open sqlsrn,"DSN=gl1197;UID=gl1197;PWD=FDT24tccG;"

response.write "<p>Open sqlsrn ok<p>"

rc=0
	while not rs.eof
		rc=rc+1
		rs.movenext
	wend

	if rc<>0 then
		ec=ec+1
		response.write "SRN already exist<p>"
	end if

	if ec=0 then
		
		response.write "Begin Posting...<p>"

		set cn=server.createobject("ADODB.connection")

		cn.open "gl1197","gl1197","FDT24tccG"
		
		response.write "<p>Connection Opened<p>"
		
		for i=0 to nr-1
			sqlupdate="UPDATE glmaster SET Balance=balance+"
			sqlupdate=sqlupdate+request.form("a00"+cstr(i)+"5field")+" WHERE "
			sqlupdate=sqlupdate+"major="+request.form("a00"+cstr(i)+"1field")+" and "
			sqlupdate=sqlupdate+"minor="+request.form("a00"+cstr(i)+"2field")+" and "
			sqlupdate=sqlupdate+"sub1="+request.form("a00"+cstr(i)+"3field")+" and "
			sqlupdate=sqlupdate+"sub2="+request.form("a00"+cstr(i)+"4field")
			
			response.write(sqlupdate+"<br>")

			cn.execute sqlupdate,numa
			
			if numa=1 then
				response.write "Update OK<p>"
			else
				response.write "Update Failed<p>"
				ec=ec+1
			end if
		
		next
		
		cn.close
		
			if ec=0 then
		
				response.write "Begin Logging Transactions...<p>"

				cn.open "gl1197","gl1197","FDT24tccG"
		
				response.write "<p>Connection Opened<p>"
				
				cn.BeginTrans
				
				for i=0 to nr-1
					ad=request.form("com0"+cstr(i))
					if ad="" then
					ad="No Description"
					end if
					sqlinsert="Insert into je values("
					sqlinsert=sqlinsert+request.form("sourceref")+", "
					sqlinsert=sqlinsert+cstr(i+1)+", "
					sqlinsert=sqlinsert+request.form("a00"+cstr(i)+"1field")+", "
					sqlinsert=sqlinsert+request.form("a00"+cstr(i)+"2field")+", "
					sqlinsert=sqlinsert+request.form("a00"+cstr(i)+"3field")+", "
					sqlinsert=sqlinsert+request.form("a00"+cstr(i)+"4field")+", "
					sqlinsert=sqlinsert+chr(39)+ad+chr(39)+", "
					sqlinsert=sqlinsert+request.form("a00"+cstr(i)+"5field")+")"
			
					response.write(sqlinsert+"<br>")

					cn.execute sqlinsert,numa
			
				if numa=1 then
					response.write "Insert je OK<p>"
				else
					response.write "Insert je Failed<p>"
					ec=ec+1
				end if
		
				next
				
				if ec=0 then
					response.write "No Errors Found... Saved Transaction"
					cn.commitTrans
					%><p><input type="button" onClick="parent.location='index.htm'" value="Return to Main Menu"><%
				else
					response.write "Errors Found... Rolling Back Transaction"
					cn.rollbackTrans 
					%><p><input type="button" onClick="parent.location='index.htm'" value="Return to Main Menu"><%
				end if
				
			else
			response.write "Errors Found During Posting... Transaction Ended<p>"
			%><p><input type="button" onClick="parent.location='index.htm'" value="Return to Main Menu"><%
			end if
	else
		response.write "Erros Found in the Voucher...  Transaction Ended<p>"
		%><p><input type="button" onClick="parent.location='index.htm'" value="Return to Main Menu"><%
	end if
%>
</body>
</html>