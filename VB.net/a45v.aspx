<html>
<body>
<p>Alan Wong
<p>0847623
<p>gl1197
<%@Page Language="VB"%>
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

<script language="VB" runat="server">

	Sub Page_Load (sender as Object, e as EventArgs)
	    Dim c as Integer=0
'		// declare a string to hold the results as an HTML table
            Dim displayoutput as String = "<center><table border='1'>"
            displayoutput=displayoutput+"<tr>"
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>MAJOR</td>"
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>MINOR</td>"
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>SUB1</td>"
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>SUB2</td>"
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>ACCOUNT DESCRIPTION</td>"
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>BALANCE</td>"
            Dim sum as Decimal=0.0
            Dim	 strConnect as String = "server=AUCKLAND;uid=gl1197;pwd=FDT24tccG;database=gl1197"

        Dim strSelect As String = "SELECT * FROM glmaster WHERE major = " & mjr.Text & _
            " AND minor = " & mnr.Text & _
            " AND sub1 = " & s1.Text & _
            " AND sub2 = " & s2.Text
        
        outSelect.InnerText = strSelect
            Try
		
			Dim  objConnect as new SqlConnection(strConnect)

'			// open the connection to the database
			objConnect.Open()

'			// create a new Command using the connection object and select statement
			Dim objCommand as new SqlCommand(strSelect, objConnect)

'			// declare a variable to hold a DataReader object
			Dim objDataReader as SqlDataReader 

'			// execute the SQL statement against the command to fill the DataReader
			objDataReader = objCommand.ExecuteReader()
'			// iterate through the records in the DataReader getting field values
'			// the Read method returns False when there are no more records
			Do while objDataReader.Read()
			        c=c+1
				displayoutput += "<tr><td>" + objDataReader("major").ToString + "</td><td>"+ objDataReader("minor").ToString + "</td><td>" _
				+ objDataReader("sub1").ToString + "</td><td>" _
				+ objDataReader("sub2").ToString + "</td><td>" _
				+ objDataReader("acctdesc") + "</td><td align='right'>" _
                                + objDataReader("balance").ToString + "</td></tr>"
                               sum=sum+Convert.ToDecimal(objDataReader("balance"))
                        Loop

'			// close the DataReader and Connection
			objDataReader.Close()
			objConnect.Close()
		
		Catch objError as Exception 
		
'			// display error details
			outError.InnerHtml = "<b>* Error </b>.<br />"+ objError.Message + "<br />" + objError.Source

		end Try

'		// add closing table tag and display the results
		displayoutput += "</table><p>Sum of Balances="
		displayoutput += sum.ToString
		displayoutput += "<p>Row Count="+c.ToString
                displayoutput += "<P>Normal Termination "+Now.toString


		outResult.InnerHtml = displayoutput

end sub	
	
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



