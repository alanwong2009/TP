<%@ Page Language="VB" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script runat="server">

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
    
    Public Sub DataGrid_Delete(Source As Object, E As DataGridCommandEventArgs)
      Dim myConnection As SqlConnection
      Dim myCommand As SqlCommand
      Dim strDeleteStmt As String
      
      Dim balance_test As Decimal
      Dim Balance_output As String
      
      balance_test=Convert.ToDecimal(E.Item.Cells(6).Text)
      
      if balance_test = 0.0 then
            outError.InnerHtml=""
            strDeleteStmt = "DELETE FROM glmaster " & _
            " WHERE major = " & E.Item.Cells(1).Text  & _
            " AND minor = " & E.Item.Cells(2).Text & _
            " AND sub1 = " & E.Item.Cells(3).Text & _
            " AND sub2 = " & E.Item.Cells(4).Text
          try
            myConnection = New SqlConnection( _
            "server=AUCKLAND;database=gl1197;uid=gl1197;pwd=FDT24tccG;")
            myCommand = New SqlCommand(strDeleteStmt, myConnection)
            myConnection.Open()
            myCommand.ExecuteNonQuery()
    
            glmasterInfo.EditItemIndex = -1
            BindData()
          Catch exc as exception
             outError.InnerHtml = "<b>* Error while deleting data</b>.<br />" _
             + exc.Message + "<br />" + exc.Source
          Finally
             myConnection.Close()
          End try
      else
          balance_output=balance_test.ToString("C")
          Dim Error_String As String
          
          Error_string="<center>You Can't Delete This Account<br>"
          Error_String=Error_String & "Balance is NOT zero.<br>"
          Error_String=Error_String & "The Account Balance is " & balance_output
          Error_String=Error_String & " for "
          Error_String=Error_String & E.Item.Cells(1).Text & "-"
          Error_String=Error_String & E.Item.Cells(2).Text & "-"
          Error_String=Error_String & E.Item.Cells(3).Text & "-"
          Error_String=Error_String & E.Item.Cells(4).Text & "</center>"
          
          outError.InnerHtml = Error_string
      end if
    End Sub
    
    Sub DataGrid_ItemCreated(Sender As Object, e As DataGridItemEventArgs)
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
		      Dim myTableCell As TableCell
		      myTableCell = e.Item.Cells(0)
	          Dim myDeleteButton As LinkButton
	          myDeleteButton = myTableCell.Controls(0)
 	          myDeleteButton.Attributes.Add("onclick","javascript: return confirm('Are you Sure you want to delete this account?');")
        End Select
    End Sub

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
      face="COMIC SANS MS"><FONT size=4><B>Delete Using a Data Grid in VB.NET<br>
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