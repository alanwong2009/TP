<html>
<%
sub pass1
%>
<body  bgcolor="#808080" link="#ffffff" text="#ffffff" vlink="#ffffff">
<p>Alan Wong
<p>0847623
<p>gl1197
<table border="2" width="95%" cellspacing="0" cellpadding="0">
<tr>
<td colspan="6">
<center>
 <br><font face="COMIC SANS MS,HELVETICA,ARIAL" size="3">    
Captain Webb's Explorations Ltd.
<br>
Update GL ACCOUNT
</center> 
</td>
</tr>

<tr>
<td>    <font face="COMIC SANS MS,HELVETICA,ARIAL" size="2">    
MAJOR
</td>
<td> <font face="COMIC SANS MS,HELVETICA,ARIAL" size="2">       
MINOR
</td>
<td> <font face="COMIC SANS MS,HELVETICA,ARIAL" size="2">       
SUB 1
</td>
<td> <font face="COMIC SANS MS,HELVETICA,ARIAL" size="2">       
SUB 2
</td>
<td><font face="COMIC SANS MS,HELVETICA,ARIAL" size="2">    
ACCOUNT DESCRIPTION
</td>
<td><font face="COMIC SANS MS,HELVETICA,ARIAL" size="2">    
BALANCE
</td>

</tr>

<!'**************** Get Delete Data -- 1st HTML>
<tr>
<td>
<FORM METHOD="POST" ACTION="a42.asp" name="f1"> 
<INPUT TYPE ="TEXT" NAME="major" size="4" maxlength="4"></td><td>
<INPUT TYPE ="TEXT" NAME="minor" size="4" maxlength="4"></td><td>
<INPUT TYPE ="TEXT" NAME="sub1" size="4" maxlength="4"></td><td>
<INPUT TYPE ="TEXT" NAME="sub2" size="4" maxlength="4"></td><td>
</td><td></td></tr>
<tr>
<td colspan="6">
<br>
<center>
<input type="hidden" name="token" value="2">
<input type="SUBMIT" value="Check if Account exist">
</FORM>
<br>
<br>
</center>
</td>
</tr>

</table>
</BODY>

</form>
<%
end sub
sub pass2

Response.write "<p>Begin...<p>"

set rs=server.createobject("ADODB.recordset")

sql="SELECT * FROM glmaster WHERE major="+request.form("major")
sql=sql+" AND minor="+request.form("minor")
sql=sql+" AND sub1="+request.form("sub1")
sql=sql+" AND sub2="+request.form("sub2")

response.write "<P>SQL="+sql+"<p>"

rs.open sql,"DSN=gl1197;UID=gl1197;PWD=FDT24tccG;"

c=0
while NOT rs.eof
    ad=rs("acctdesc")
    c=c+1
    rs.movenext
wend
rs.close
set rs=nothing

if c=0 then
  response.write "ACCT NOT FOUND"
 %>
 <br><input type="button" onClick="parent.location='index.htm'" value="Return to Main Menu">
 <%
else
  response.write "ACCT FOUND"
%>
<form action="a42.asp" method="POST" name="f1">
<br><input type="hidden" name="major" value="<% =request.form("major") %>"> MAJOR=<% =request.form("major") %>
<br><input type="hidden" name="minor" value="<% =request.form("minor") %>"> MINOR=<% =request.form("minor") %>
<br><input type="hidden" name="sub1" value="<% =request.form("sub1") %>"> SUB1=<% =request.form("sub1") %>
<br><input type="hidden" name="sub2" value="<% =request.form("sub2") %>"> SUB2=<% =request.form("sub2") %>
<br>Modify this description :<input type="text" size="50" value="<% = ad %>" name="ad">
<br><input type="hidden" name="token" value="3">
<br><input type="SUBMIT" value="submit">
</form>

<%
end if

end sub

sub pass3

set cn=server.createobject("ADODB.connection")
cn.open "gl1197","gl1197","FDT24tccG"

sql ="UPDATE glmaster SET acctdesc="+chr(39)+Request.form("ad")+chr(39)
sql=sql+" WHERE major="+request.form("major") +" AND "
sql=sql+" minor="+request.form("minor") +" AND "
sql=sql+" sub1="+request.form("sub1") +" AND "
sql=sql+" sub2="+request.form("sub2")

response.write "<p>"+sql+"<P>"

cn.execute sql,numa

if numa=1 then
 response.write "<p>UPDATE OK"
%>
<br><input type="button" onClick="parent.location='index.htm'" value="Return to Main Menu">
<%
else
 response.write "<p>UPDATE FAILED"
%>
 <br><input type="button" onClick="parent.location='index.htm'" value="Return to Main Menu">
<%
end if

cn.close
set cn=nothing

end sub
'
'*****MAIN
'
token_value=Request.form("token")
select case token_value
case ""
  call pass1
case "2"
  call pass2
case "3"
  call pass3
end select
%>
</body>
</html>