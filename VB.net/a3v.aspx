<html>
<body>
<p>Alan Wong
<p>0847623
<p>gl1197
<%@Page Language="VB"%>
<%@Import Namespace="System.Data" %>
<%@Import Namespace="System.Data.SqlClient" %>

<div>SELECT command: <b><span id="outSelect" runat="server"></span></b></div>
<div id="select1" runat="server"> </div>
<div>SRN command: <b><span id="outsrn" runat="server"></span></b></div>
<div id="srn" runat="server"> </div>
<div>Update command: <b><span id="outupdate" runat="server"></span></b></div>
<div id="update" runat="server"></div>
<div>Insert command: <b><span id="outinsert" runat="server"></span></b></div>
<div id="insert" runat="server"></div>
<p><input type="button" onClick="parent.location='index.htm'" value="Return to Main Menu">
<script language="VB" runat="server">
    Sub Page_Load(sender As Object, e As EventArgs)
        Console.WriteLine("Begin...")
    
        Dim ec As Integer = 0
        Dim nr As Integer = 0
        Dim poop As String
        For i As Integer = 0 To 6
            poop = Request.Form("a00" + CStr(i) + "1field")
            If poop <> "" Then
                nr = nr + 1
            End If
        Next
        
        Dim Sql As String
        Dim login As String = "server=AUCKLAND;uid=gl1197;pwd=FDT24tccG;database=gl1197"
        Dim rc As Integer = 0
        Try
            Dim objConnect As New SqlConnection(login)
            objConnect.Open()
            
            For j As Integer = 0 To nr - 1
                Sql = "select * from glmaster where "
                Sql = Sql + "major=" + Request.Form("a00" + CStr(j) + "1field") + " and "
                Sql = Sql + "minor=" + Request.Form("a00" + CStr(j) + "2field") + " and "
                Sql = Sql + "sub1=" + Request.Form("a00" + CStr(j) + "3field") + " and "
                Sql = Sql + "sub2=" + Request.Form("a00" + CStr(j) + "4field")
                outSelect.InnerText = Sql
                Dim SqlCommand As New SqlCommand(Sql, objConnect)
                Dim SqlDataReader As SqlDataReader
                SqlDataReader = SqlCommand.ExecuteReader()
                
                Do While SqlDataReader.Read()
                    rc = rc + 1
                Loop
                SqlDataReader.Close()
            Next
            
            objConnect.Close()
        Catch select1Error As Exception
            ec = ec + 1
            select1.InnerHtml = "<b>* Error </b>.<br />" + select1Error.Message + "<br />" + select1Error.Source
        End Try
        
        If rc <> nr Then
            ec = ec + 1
            select1.InnerHtml = "Account Not Found!<p>"
        End If
        
        If ec = 0 Then
            rc = 0
            Dim sqlsrn As String = "select * from je where sourceref=" + Request.Form("sourceref")
            outsrn.InnerHtml = sqlsrn
                
            Try
                Dim objConnect As New SqlConnection(login)
                objConnect.Open()
                Dim srncommand As New SqlCommand(sqlsrn, objConnect)
                Dim srndatareader As SqlDataReader
                srndatareader = srncommand.ExecuteReader()
                Do While srndatareader.Read()
                    rc = rc + 1
                Loop
                srndatareader.Close()
                objConnect.Close()
            Catch srnerror As Exception
                ec = ec + 1
                srn.InnerHtml = "<b>* Error </b>.<br />" + srnerror.Message + "<br />" + srnerror.Source
            End Try
                 
            If rc <> 0 Then
                ec = ec + 1
                srn.InnerHtml = "SRN already exist<p>"
            End If
                 
            If ec = 0 Then
                Dim sqlupdate As String
                             
                For i As Integer = 0 To nr - 1
                    sqlupdate = "UPDATE glmaster SET Balance=balance+"
                    sqlupdate = sqlupdate + Request.Form("a00" + CStr(i) + "5field") + " WHERE "
                    sqlupdate = sqlupdate + "major=" + Request.Form("a00" + CStr(i) + "1field") + " and "
                    sqlupdate = sqlupdate + "minor=" + Request.Form("a00" + CStr(i) + "2field") + " and "
                    sqlupdate = sqlupdate + "sub1=" + Request.Form("a00" + CStr(i) + "3field") + " and "
                    sqlupdate = sqlupdate + "sub2=" + Request.Form("a00" + CStr(i) + "4field")
                                          
                    outupdate.InnerHtml = sqlupdate
                                          
                    Try
                        Dim objConnect As New SqlConnection(login)
                        Dim updatecommand As New SqlCommand(sqlupdate, objConnect)
                        objConnect.Open()
                        updatecommand.ExecuteNonQuery()
                        objConnect.Close()
                    Catch updateerror As Exception
                        ec = ec + 1
                        update.InnerHtml = "<b>* Error </b>.<br />" + updateerror.Message + "<br />" + updateerror.Source
                    End Try                    
                Next
                
                If ec = 0 Then
                    Dim sqlinsert As String
                    Dim ad As String
                    For i As Integer = 0 To nr - 1
                        ad = Request.Form("com0" + CStr(i))
                        If ad = "" Then
                            ad = "No Description"
                        End If
                        sqlinsert = "Insert into je values("
                        sqlinsert = sqlinsert + Request.Form("sourceref") + ", "
                        sqlinsert = sqlinsert + CStr(i + 1) + ", "
                        sqlinsert = sqlinsert + Request.Form("a00" + CStr(i) + "1field") + ", "
                        sqlinsert = sqlinsert + Request.Form("a00" + CStr(i) + "2field") + ", "
                        sqlinsert = sqlinsert + Request.Form("a00" + CStr(i) + "3field") + ", "
                        sqlinsert = sqlinsert + Request.Form("a00" + CStr(i) + "4field") + ", "
                        sqlinsert = sqlinsert + Chr(39) + ad + Chr(39) + ", "
                        sqlinsert = sqlinsert + Request.Form("a00" + CStr(i) + "5field") + ")"
			
                        outinsert.InnerHtml = sqlinsert + "<br>"
                        
                        Try
                            Dim objConnect As New SqlConnection(login)
                            Dim insertcommand As New SqlCommand(sqlinsert, objConnect)
                            objConnect.Open()
                            insertcommand.ExecuteNonQuery()
                            objConnect.Close()
                        Catch inserterror As Exception
                            ec = ec + 1
                            insert.InnerHtml = "<b>* Error </b>.<br />" + inserterror.Message + "<br />" + inserterror.Source
                        End Try
                        
                    Next
                    
                End If
            End If
        End If
    End Sub
</script>
</body>
</html>
