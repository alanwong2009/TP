<html>
<HEAD>
<SCRIPT LANGUAGE="VBScript">
<!--
Sub GoBackBut_OnClick

window.location.href="http://disc-nt.cba.uh.edu/Students/gl1197/"

end sub
-->
</script>
</HEAD>
<body  bgcolor="#cccccc" link="#ffffff" text="#000000" vlink="#ffffff">
<p>Alan Wong
<p>0847623
<p>gl1197
<table border="1"><tr>
<td valign="middle" align="center" bgcolor="#ffffff"><font size='4'>
<b> RESTORING YOUR <i>glmaster</i> and <i>je</i> TABLES 

</td></tr></table>
<p>
<font color="#000000">
<p><b>Restoring the General Ledger Tables:</b><p>
1. Open SRCGL
<br>    (the permanent copy of the general ledger database)

<% '*************************** functions, subs, then main

function fixdec (amount)
dim la, lb, dif,temp
  temp=cStr(amount)
  la=InStr(1,temp,".")
  lb=len(temp)
  dif=lb-la
  if la = 0 then
      temp=temp+".00"
  else
      if  dif=0 then
          temp=temp+"00"
     else
          if dif=1 then
                temp=temp+"0"
          else

          end if
    end if
 end if
 fixdec=cStr(temp)
end function

sub buildglmaster (cnnew)
dim create_string
on error resume next
   create_string="CREATE TABLE glmaster ("
   create_string=create_string + "major integer NOT NULL,"
   create_string=create_string + "minor integer NOT NULL,"
   create_string=create_string + "sub1 integer,"
   create_string=create_string + "sub2 integer,"
   create_string=create_string + "acctdesc varchar(50),"
   create_string=create_string + "balance numeric (18,2) NOT NULL "
   create_string=create_string + "primary key (major, minor, sub1, sub2) )"
   cnnew.execute create_string
   if noerrors(cnnew, "Task: Create new glmaster table") then
         Response.write "<br>4. Created new glmaster table OK"
   else
         Response.write "<br>4. Create new glmaster table task failed ******<br>"
   end if
end sub

sub dropglmaster (cnnew)
on error resume next
   cnnew.execute "DROP TABLE glmaster", numa
   if noerrors(cnnew, "Task: drop glmaster table") then
      Response.write "<br>3. Dropped old glmaster table OK"
   else
      Response.write "<br>3. Unable to drop glmaster table. Task Failed ****<br>"
   end if
   buildglmaster (cnnew)
end sub

sub dropje (cnnew)
on error resume next
    cnnew.execute "DROP TABLE je", numa
    if noerrors(cnnew, "Task: dropping je table") then
       Response.write "<p>5. Dropped old je table OK"
    else
       Response.write "<p>5. Unable to  drop je table. Task Failed *********<br>"
    end if
    buildje(cnnew)
end sub

sub buildje (cnnew)
dim create_string
on error resume next
   create_string="CREATE TABLE je ("
   create_string=create_string +"sourceref integer NOT NULL,"
   create_string=create_string +"srseq integer NOT NULL,"
   create_string=create_string +"jemajor integer NOT NULL,"
   create_string=create_string +"jeminor integer NOT NULL,"
   create_string=create_string +"jesub1 integer,"
   create_string=create_string +"jesub2 integer,"
   create_string=create_string +"jedesc char(50),"
   create_string=create_string +"jeamount numeric (18,2) NOT NULL "
   create_string=create_string +"primary key (sourceref, srseq) )"
   cnnew.execute create_string
   if noerrors(cnnew, "Task: Create new je table") then
         Response.write "<br>6. Created new je table OK"
   else
         Response.write "<br>6. je table create task failed ***************<br>"
   end if
end sub

Function noerrors (cn , task)
If Err <> 0 Then
    If cn.Errors.Count = 0 Then

    Else
         for i = 0 to cn.Errors.Count - 1
              response.write "<p>"
              response.write cn.errors(i)
         next
    End If
    noerrors= False
Else
    noerrors = True
End if
End Function

'*************************** main ****************************************

dim cm,cnold,cnnew
dim sumbal, rsold, Insert_String,create_string,numa,numnew,deval
dim fdsn, fuid,fpwd
sumbal=0
numnew=0
on error resume next

'************ user's DSN, Userid and Password ****************************
'*                                                                       *
'*                                                                       *
'*            CHANGE THE THREE VALUES BELOW                              *
'*            TO THE DATABASE CREDENTIALS PROVIDED                       *
'*            TO YOU IN CLASS                                            *
'*                                                                       *
'*                                                                       *
'*************************************************************************

fdsn="gl1197"
fuid="gl1197"
fpwd="FDT24tccG"


'*********** open the original general ledger database

set rsold = Server.CreateObject("ADODB.Recordset")
oldsql="SELECT * FROM glmaster order by major ASC, minor ASC, sub1 ASC, sub2 ASC"
Response.write "<br>   With sql="+oldsql
rsold.open oldsql,"DSN=SRCGL;UID=;PWD=;"
Response.write "<br>   Opened SRCGL OK"

'*********** open the user requested user database

set cnnew = Server.CreateObject("ADODB.Connection")
cnnew.open fdsn,fuid,fpwd

if noerrors (cnnew, "Task: Opening database") then '******** top test for user database
          Response.write "<br>2. Opened your "
          Response.write  "<b>"+fdsn+"</b>"
          Response.write " database OK"

          call dropglmaster(cnnew) '****** drop, then create the glmaster table

          Response.write "<p><table border="+chr(34)+"1"+chr(34)+">"
          Response.write "<tr><td>major</td><td>minor</td><td>sub1</td><td>sub2</td><td>"
          Response.write "acctdesc</td><td>balance</td></tr>"

          while not rsold.EOF   '****** loop thru the SRCGL table rows,
'                                                        copying each  to the new glmaster

                 Insert_String = "INSERT INTO glmaster (major,minor,sub1,sub2,acctdesc,balance) VALUES ("
                 Insert_String = Insert_String + cStr(rsold("major")) + ","
                 Insert_String = Insert_String + cStr(rsold("minor")) + ","
                 Insert_String = Insert_String + cStr(rsold("sub1")) + ","
                 Insert_String = Insert_String + cStr(rsold("sub2")) + ","
                 Insert_String = Insert_String + chr(39)+cStr(rsold("acctdesc")) + chr(39) + ","
                 Insert_String = Insert_String + cStr(rsold("balance")) + ")"

                 cnnew.execute Insert_String '****** add the row to the new table

                 Response.write "<tr><td align="+chr(34)+"right"+chr(34)+">" '************ show the user
                 Response.write rsold("major")
                 Response.write "</td><td align="+chr(34)+"right"+chr(34)+">"
                 Response.write rsold("minor")
                 Response.write "</td><td align="+chr(34)+"right"+chr(34)+">"
                 Response.write rsold("sub1")
                 Response.write "</td><td align="+chr(34)+"right"+chr(34)+">"
                 Response.write rsold("sub2")
                 Response.write "</td><td align="+chr(34)+"right"+chr(34)+">"
                 Response.write rsold("acctdesc")
                 Response.write "</td><td align="+chr(34)+"right"+chr(34)+">"

                 deval=rsold("balance") '************* get the balance
                 deval=fixdec(deval)

                 Response.write deval '************** write the balnce to the user
                 Response.write "</td></tr>"

                 sumbal=sumbal+cDbl(rsold("balance")) '************* count the rows
                 numnew=numnew+1

                 rsold.movenext '************  get next record from SRCGL

             wend '**************************  end of add rows loop
             rsold.close '*******************  finished with the new glmaster table, so close the SRCGL version

            Response.write "<tr><td colspan="+chr(34)+"5" +chr(34)+"align="
            Response.write chr(34)+"right"+chr(34)+">SUM OF BALANCES</td><td align="+chr(34)+"right"+chr(34)+">"
            Response.write cStr(fixdec(sumbal))
            Response.write "</td></tr><tr><td colspan="+chr(34)+"6"+chr(34)+">"
            Response.write cStr(numnew)
            Response.write " rows in the new glmaster table</td></tr></table><br>"

           call dropje(cnnew) '***** now drop, then create the je table

else '**************************** couldn't open the user's database -- bail out!

       Response.write "<p>2. Open task failed on database: "
       Response.write cStr(fdsn)
       Response.write "<p>3. Drop current glmaster table NOT attempted"
       Response.write "<p>4. Create glmaster table NOT attempted"
       Response.write "<p>5. Drop current je table NOT attempted"
       Response.write "<p>6. Create je table NOT attempted"

end if

'**************************** back to HTML for the finish

cnnew.close
set cnnew=nothing
set rsold=nothing

%>

<br>
<p>
<center><b>END OF THE GENERAL LEDGER RESTORE</center>
<p>
<form>
<input type="button" value="Back to Main Menu" name="GoBackBut">
</form>
</body>
</html>