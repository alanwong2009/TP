<html>
<body>
<p>Alan Wong
<p>0847623
<p>gl1197
<%@Page Language="C#"%>
<%@Import Namespace="System.Data" %>
<%@Import Namespace="System.Data.SqlClient" %>
<center><table border="1">
<tr>
<td><img src="captsm.gif"></td>
<td valign="middle" align="center" bgcolor="#aaaaaa"><font face="COMIC SANS MS"><font size="4">
<b>The Simple Table Listing Program Using a Data Grid<br>in VB.NET<hr>Lists the Contents of <i>glmaster</i><br>Using a DataReader in SQL
</td>
</tr></table>

<div>SELECT command: <b><span id="outSelect" runat="server"></span></b></div>
<div id="outError" runat="server"> </div>
<div id="outResult" runat="server"></div>

<script language="C#" runat="server">
    
	void Page_Load (Object sender, EventArgs e)
    {
	    int c = 0;
            String displayoutput = "<center><table border='1'>";
            displayoutput=displayoutput+"<tr>";
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>MAJOR</td>";
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>MINOR</td>";
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>SUB1</td>";
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>SUB2</td>";
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>ACCOUNT DESCRIPTION</td>";
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>BALANCE</td>";
            Decimal sum = (decimal)0.0;
            String strConnect = "server=AUCKLAND;uid=gl1197;pwd=FDT24tccG;database=gl1197";

        String strSelect = "SELECT * FROM glmaster WHERE major = " + mjr.Text + " AND minor = " + mnr.Text + " AND sub1 = "+ s1.Text + " AND sub2 = " + s2.Text;
        
        outSelect.InnerText = strSelect;
            try
		{
			SqlConnection objConnect = new SqlConnection(strConnect);
			objConnect.Open();
			SqlCommand objCommand = new SqlCommand(strSelect, objConnect);
			SqlDataReader objDataReader;
			objDataReader = objCommand.ExecuteReader();
			while (objDataReader.Read())
			{
			  displayoutput += "<tr><td>" + objDataReader["major"] + "</td><td>"
          				 + objDataReader["minor"] + "</td><td>"
				         + objDataReader["sub1"] + "</td><td>"
				         + objDataReader["sub2"] + "</td><td>"
				         + objDataReader["acctdesc"] + "</td><td align='right'>"
			                 + objDataReader["balance"] + "</td></tr>";
                          sum=sum+Convert.ToDecimal(objDataReader["balance"]);						  
                          c++;
			}
			objDataReader.Close();
			objConnect.Close();
        }
		catch (Exception objError)
            {
			    outError.InnerHtml = "<b>* Error </b>.<br />"+ objError.Message + "<br />" + objError.Source;
            }
		displayoutput += "</table><p>Sum of Balances=";
		displayoutput += sum;
		displayoutput += "<p>Row Count="+c.ToString();
		DateTime t = DateTime.Now;
                displayoutput += "<P>Normal Termination "+t;

		outResult.InnerHtml = displayoutput;

}
	
</script>
<form id="Form1" method="post" runat="server">
<p>Major <asp:TextBox columns="4" ID="mjr" Runat="server" />
<asp:RequiredFieldValidator id="RequiredFieldValidator1" runat="server" EnableClientScript="False" Width="208px" Display="Dynamic" ControlToValidate="mjr" ErrorMessage="Major Account Code is Required"></asp:RequiredFieldValidator>  
<asp:RangeValidator id="RangeValidator1" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="mjr" ErrorMessage="Major Account must be between 0 and  9999" MinimumValue="0" MaximumValue="9999" Type="Integer"></asp:RangeValidator>   
<asp:CompareValidator id="CompareValidator1" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="mjr" ErrorMessage="Must be digits 0-9 only" Operator="DataTypeCheck" Type="Integer"></asp:CompareValidator> <br />
<br>Minor <asp:TextBox columns="4" ID="mnr" Runat="server" />
<asp:RequiredFieldValidator id="RequiredFieldValidator2" runat="server" EnableClientScript="False" Width="208px" Display="Dynamic" ControlToValidate="mnr" ErrorMessage="Major Account Code is Required"></asp:RequiredFieldValidator>  
<asp:RangeValidator id="RangeValidator2" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="mnr" ErrorMessage="Major Account must be between 0 and  9999" MinimumValue="0" MaximumValue="9999" Type="Integer"></asp:RangeValidator>   
<asp:CompareValidator id="CompareValidator2" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="mnr" ErrorMessage="Must be digits 0-9 only" Operator="DataTypeCheck" Type="Integer"></asp:CompareValidator> <br />
<br>Sub 1 <asp:TextBox columns="4" ID="s1" Runat="server" />
<asp:RequiredFieldValidator id="RequiredFieldValidator3" runat="server" EnableClientScript="False" Width="208px" Display="Dynamic" ControlToValidate="s1" ErrorMessage="Major Account Code is Required"></asp:RequiredFieldValidator>  
<asp:RangeValidator id="RangeValidator3" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="s1" ErrorMessage="Major Account must be between 0 and  9999" MinimumValue="0" MaximumValue="9999" Type="Integer"></asp:RangeValidator>   
<asp:CompareValidator id="CompareValidator3" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="s1" ErrorMessage="Must be digits 0-9 only" Operator="DataTypeCheck" Type="Integer"></asp:CompareValidator> <br />
<br>Sub 2 <asp:TextBox columns="4" ID="s2" Runat="server" />
<asp:RequiredFieldValidator id="RequiredFieldValidator4" runat="server" EnableClientScript="False" Width="208px" Display="Dynamic" ControlToValidate="s2" ErrorMessage="Major Account Code is Required"></asp:RequiredFieldValidator>  
<asp:RangeValidator id="RangeValidator4" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="s2" ErrorMessage="Major Account must be between 0 and  9999" MinimumValue="0" MaximumValue="9999" Type="Integer"></asp:RangeValidator>   
<asp:CompareValidator id="CompareValidator4" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="s2" ErrorMessage="Must be digits 0-9 only" Operator="DataTypeCheck" Type="Integer"></asp:CompareValidator> <br />
<p><asp:Button Text="Query Account" ID="btnquery" OnClick="Page_Load" Runat="server" />
<p><asp:Label ID="numa" ForeColor="#ff0000" Runat=Server/>
</form>
</body>
</html>



