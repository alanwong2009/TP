<%@ Page Language="VB" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    '
    'NOTE:
    '
    'In the HTML below is a'template' column (as opposed to a 'bound' column)
    'for the textbox (id="adtextbox").
    '
    'It gets its value from the DataBinder.Eval
    '
    '  A RequiredFieldValidator and a RegularExpressionValidator are used to test
    '  the validity of the "adtextbox"
    '
    '
    
         Public Sub Page_Load(Source As Object, E As EventArgs)
           If Not Page.IsPostBack Then
             BindData()
           End If
         End Sub
    
         Public Sub BindData()
           Dim myDataSet As New DataSet
           Dim mySqlDataAdapter As SqlDataAdapter
           mySqlDataAdapter = New SqlDataAdapter("SELECT * FROM glmaster","server=AUCKLAND;database=gl1197;uid=gl1197;pwd=FDT24tccG;")
           mySqlDataAdapter.Fill(myDataSet, "glmaster")
           glmasterInfo.DataSource = myDataSet.Tables("glmaster")
           glmasterInfo.DataBind()
         End Sub
    
         Public Sub DataGrid_Edit(Source As Object, E As DataGridCommandEventArgs)
           glmasterInfo.EditItemIndex = E.Item.ItemIndex
           BindData()
         End Sub
    
         Public Sub DataGrid_Cancel(Source As Object, E As DataGridCommandEventArgs)
           glmasterInfo.EditItemIndex = -1
           BindData()
         End Sub
    
         Public Sub DataGrid_Update(Source As Object, E As DataGridCommandEventArgs)
           Dim myConnection As SqlConnection
           Dim myCommand As SqlCommand
           Dim strUpdateStmt As String
         '
         'Rather than retreiveing the account description from a bound column
         '  (i.e., E.Item.Cells(column_number).Text, the acctdesc value is retreived by
         '   finding the value in the template column (E.Item.FindControl(adtextbox))
         '   and converting its value (Server.HtmlEncode)
    
           Dim acctdesctextbox As TextBox
           acctdesctextbox = E.Item.FindControl("adtextbox")
           Dim acctdescfound As String = Server.HtmlEncode(acctdesctextbox.Text)
    
           strUpdateStmt = "UPDATE glmaster SET " & _
            "acctdesc = '" & acctdescfound & "'" & _
            " WHERE major = " & E.Item.Cells(1).Text  & _
            " AND minor = " & E.Item.Cells(2).Text & _
            " AND sub1 = " & E.Item.Cells(3).Text & _
            " AND sub2 = " & E.Item.Cells(4).Text
         '
         '  The page is checked for Validity (Page.IsValid) since data errors
         '   should abort the update
         '
           if Page.IsValid then
    
           try
             myConnection = New SqlConnection( _
              "server=AUCKLAND;database=gl1197;uid=gl1197;pwd=FDT24tccG;")
             myCommand = New SqlCommand(strUpdateStmt, myConnection)
             myConnection.Open()
             myCommand.ExecuteNonQuery()
    
             glmasterInfo.EditItemIndex = -1
             BindData()
            Catch exc As Exception
                outError.InnerHtml = "<b>* Error while updating data</b>.<br />" _
                + exc.Message + "<br />" + exc.Source
           finally
              myConnection.Close()
           End try
    
           end if
    
         End Sub

</script>
<html>
<head>
</head>
<body>
<p>Alan Wong
<p>0847623
<p>gl1197
    <center>
        <table border="1">
            <tbody>
                <tr>
                    <td>
                        <img src="captsm.gif" /></td>
                    <td valign="center" align="middle" bgcolor="#aaaaaa">
                        <font face="COMIC SANS MS"><font size="4"><b>Update (Modify) Using a Data Grid<br />
                        in VB.NET 
                        <hr />
                        Changing the Account Description for the <i>glmaster</i> </b></font></font></td>
                </tr>
            </tbody>
        </table>
        <form method="post" runat="server">
            <div id="outError" runat="server">
            </div>
            <br />
            <asp:DataGrid id="glmasterInfo" runat="server" AutoGenerateColumns="False" OnEditCommand="DataGrid_Edit" OnCancelCommand="DataGrid_Cancel" OnUpdateCommand="DataGrid_Update">
                <Columns>
                    <asp:EditCommandColumn HeaderText="<b>USER<br>ACTION " ButtonType="LinkButton" CancelText="Cancel" EditText="Edit" UpdateText="Update" />
                    <asp:BoundColumn DataField="major" HeaderText="<b>MAJOR" ReadOnly="True" />
                    <asp:BoundColumn DataField="minor" HeaderText="<b>MINOR" ReadOnly="True" />
                    <asp:BoundColumn DataField="sub1" HeaderText="<b>SUB<br>ACCOUNT<br>1" ReadOnly="True" />
                    <asp:BoundColumn DataField="sub2" HeaderText="<b>SUB<br>ACCOUNT<br>2" ReadOnly="True" />
                    <asp:TemplateColumn HeaderText="<b>Account<br>Description">
                        <ItemTemplate>
                            <asp:Label runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "acctdesc") %>' />
                        </ItemTemplate>
                        <EditItemTemplate>
                            <nobr /> 
                            <asp:TextBox runat="server" id="adtextbox" Text='<%# DataBinder.Eval(Container.DataItem, "acctdesc") %>' />
                            <asp:RequiredFieldValidator id="adreqd" ControlToValidate="adtextbox" Display="Dynamic" runat="server">
                     Account decsription cannot be blank
            </asp:RequiredFieldValidator>
            <!-- 
              *************************CAUTION *******************************
              the following line has an oddity in it. In order prevent the user
              from entering a single or double quote the regular expression validator uses the pattern: [^\'\"]+
              BUT the validator pattern is enclosed in double quote marks itself.
              So, to prevent the double quote from terminating the ValidationExpression string
              We have to force the double quote to be inserted at execution time.
              So in the line below, replace the double quote with the string: AMPERSANDCHARACTERquote;
              that is, the single ampersand character followed by the string  "quot;"
              The following line is done this way. But you can't see it when you view this page,
              because the ampersandquot; has been rendered by your browser. But it is there nonetheless.
              See this discussed: http://forums.asp.net/t/1220701.aspx/1
              ******************************************************************-->
              
                            
                        </EditItemTemplate>
                    </asp:TemplateColumn>
                    <asp:BoundColumn DataField="balance" HeaderText="<b>ACCOUNT<br>BALANCE" ReadOnly="True" />
                </Columns>
            </asp:DataGrid>
        </form>
    </center>
</body>
</html>