<%@ Page Language="C#" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script runat="server">   
         public void Page_Load(Object sender, EventArgs e)
         {
           if (!Page.IsPostBack)
             
             {
              BindData();
             }
           
         }
    
         public void BindData()
         {
           DataSet myDataSet = new DataSet();
           SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
           mySqlDataAdapter = new SqlDataAdapter("SELECT * FROM glmaster","server=AUCKLAND;database=gl1197;uid=gl1197;pwd=FDT24tccG;");
           mySqlDataAdapter.Fill(myDataSet, "glmaster");
           glmasterInfo.DataSource = myDataSet.Tables[0];
           glmasterInfo.DataBind();
         }
    
         public void DataGrid_Edit(Object Source , DataGridCommandEventArgs E)
         {
            glmasterInfo.EditItemIndex = E.Item.ItemIndex;
            BindData();
         }
         
    
         public void DataGrid_Cancel(Object Source ,DataGridCommandEventArgs E)
         {
           glmasterInfo.EditItemIndex = -1;
           BindData();
         }
    
         public void DataGrid_Update(Object Source ,DataGridCommandEventArgs E)
         {
           String strUpdateStmt;
           TextBox acctdesctextbox = new TextBox();
           acctdesctextbox = (TextBox)E.Item.FindControl("adtextbox");
           String acctdescfound = Server.HtmlEncode(acctdesctextbox.Text);   
           strUpdateStmt = "UPDATE glmaster SET acctdesc = '" + acctdescfound + "'" + " WHERE major = " + E.Item.Cells[1].Text  + " AND minor = "+ E.Item.Cells[2].Text + " AND sub1 = " + E.Item.Cells[3].Text +  " AND sub2 = " + E.Item.Cells[4].Text;
           
             if (Page.IsValid)
                {
                try
                    {
                    SqlConnection myConnection = new SqlConnection("server=AUCKLAND;database=gl1197;uid=gl1197;pwd=FDT24tccG;");
                    SqlCommand myCommand = new SqlCommand(strUpdateStmt, myConnection);
                    myConnection.Open();
                    myCommand.ExecuteNonQuery();
                    glmasterInfo.EditItemIndex = -1;
                    BindData();
                    myConnection.Close();
                }
                catch (Exception exc)
                    {
                    outError.InnerHtml = "<b>* Error while updating data</b>.<br />"+ exc.Message + "<br />" + exc.Source;
                    }
             }
         }    
            
                      

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
                        in C# 
                        <hr />
                        Changing the Account Description for the <i>glmaster</i> </b></font></font></td>
                </tr>
            </tbody>
        </table>
        <form id="Form1" method="post" runat="server">
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
                            <asp:Label ID="Label1" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "acctdesc") %>' />
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