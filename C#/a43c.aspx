<%@ Page Language="C#" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    public void Page_Load(Object Source, EventArgs E)
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
        mySqlDataAdapter = new SqlDataAdapter("SELECT * FROM glmaster", "server=AUCKLAND;database=gl1197;uid=gl1197;pwd=FDT24tccG;");
        mySqlDataAdapter.Fill(myDataSet, "glmaster");
        glmasterInfo.DataSource = myDataSet.Tables[0];
        glmasterInfo.DataBind();
    }
    
    public void DataGrid_Delete(Object Source, DataGridCommandEventArgs E)
    {
        String strDeleteStmt;
        Decimal balance_test;
        String Balance_output;
      
        balance_test = Convert.ToDecimal(E.Item.Cells[6].Text);
      
        if (balance_test == (decimal)0.0) 
        {
            outError.InnerHtml = "";
            strDeleteStmt = "DELETE FROM glmaster WHERE major = " + E.Item.Cells[1].Text + " AND minor = " + E.Item.Cells[2].Text + " AND sub1 = " + E.Item.Cells[3].Text + " AND sub2 = " + E.Item.Cells[4].Text;
            try
            {
                SqlConnection myConnection = new SqlConnection("server=AUCKLAND;database=gl1197;uid=gl1197;pwd=FDT24tccG;");
                SqlCommand myCommand = new SqlCommand(strDeleteStmt, myConnection);
                myConnection.Open();
                myCommand.ExecuteNonQuery();
    
                glmasterInfo.EditItemIndex = -1;
                BindData();
                myConnection.Close();
            }
                catch (Exception exc)
                 {
                     outError.InnerHtml = "<b>* Error while deleting data</b>.<br />" + exc.Message + "<br />" + exc.Source;
                 }
   
        }
        else
        {
            Balance_output = balance_test.ToString("C");
            String Error_String;
          
            Error_String = "<center>You Can't Delete This Account<br>";
            Error_String = Error_String + "Balance is NOT zero.<br>";
            Error_String = Error_String + "The Account Balance is " + Balance_output;
            Error_String = Error_String + " for ";
            Error_String = Error_String + E.Item.Cells[1].Text + "-";
            Error_String = Error_String + E.Item.Cells[2].Text + "-";
            Error_String = Error_String + E.Item.Cells[3].Text + "-";
            Error_String = Error_String + E.Item.Cells[4].Text + "</center>";
          
            outError.InnerHtml = Error_String;
        }
    }
    
    void DataGrid_ItemCreated(Object Sender, DataGridItemEventArgs e )
    {
        int caseswitch = Convert.ToInt32(e.Item.ItemType);
        switch (caseswitch)
        {
            case (int)ListItemType.Item: 
            case (int)ListItemType.AlternatingItem: 
            case (int)ListItemType.EditItem:       
                TableCell myTableCell; 
                myTableCell = e.Item.Cells[0];
                LinkButton myDeleteButton ;
                myDeleteButton = (LinkButton)myTableCell.Controls[0];
                myDeleteButton.Attributes.Add("onclick", "javascript: return confirm('Are you Sure you want to delete this account?');");
                break;
        }
    }

</script>
<html>
<BODY>
<p>Alan Wong
<p>0847623
<p>gl1197
<CENTER>
<TABLE border=1>
  <TBODY>
  <TR>
    <TD><IMG src="captsm.gif"></TD>
    <TD vAlign=center align=middle bgColor=#aaaaaa><FONT 
      face="COMIC SANS MS"><FONT size=4><B>Delete Using a Data Grid in C#<br>
      With User Confirmation 
      <HR>
      Removing a Row From <I>glmaster</I> </B></FONT></FONT></TD></TR></TBODY></TABLE>
<p>
<font color="#ff0000"><b>
<DIV id=outError runat="server"> </DIV>
</b></font>
<FORM method=post runat="server">
<asp:DataGrid id=glmasterInfo runat="server" 
 onItemCreated="DataGrid_ItemCreated"
 OnDeleteCommand="DataGrid_Delete" AutoGenerateColumns="False">
 <Columns>
  <asp:ButtonColumn
    HeaderText=" "
    Text="Delete"
    commandname="Delete"
    ButtonType="LinkButton" />
  <asp:BoundColumn
    DataField="major"
    HeaderText="<b>MAJOR"
    ReadOnly="True" />
  <asp:BoundColumn
    DataField="minor"
    HeaderText="<b>MINOR"
    ReadOnly="True" />
  <asp:BoundColumn
    DataField="sub1"
    HeaderText="<b>SUB<br>ACCOUNT<br>1"
    ReadOnly="True" />
  <asp:BoundColumn
    DataField="sub2"
    HeaderText="<b>SUB<br>ACCOUNT<br>2"
    ReadOnly="True" />
  <asp:BoundColumn
    DataField="acctdesc"
    HeaderText="<b>ACCOUNT<br>DESCRIPTION" />
  <asp:BoundColumn
    DataField="balance"
    HeaderText="<b>ACCOUNT<br>BALANCE"
    ReadOnly="True"/>
 </Columns>
</asp:DataGrid></FORM></CENTER></BODY></HTML>