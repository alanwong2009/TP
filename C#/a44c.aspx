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
<td valign="middle" align="center" bgcolor="#aaaaaa"><font face="COMIC SANS MS"><font size="4">
<b>The Simple Table Listing Program Using a Data Grid<br>in C#<hr>Lists the Contents of <i>glmaster</i><br>Using a DataReader in SQL
</td>
</tr></table>

<div>SELECT command: <b><span id="outSelect" runat="server"></span></b></div>
<div id="outError" runat="server"> </div>
<div id="outResult" runat="server"></div>

<script language="C#" runat="server">

	void Page_Load(Object sender, EventArgs e)
	{
		// declare a string to hold the results as an HTML table
            
            string displayoutput = "<center><table>";
            int c=0;
            displayoutput=displayoutput+"<tr>";
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>MAJOR</td>";
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>MINOR</td>";
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>SUB1</td>";
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>SUB2</td>";
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>ACCOUNT DESCRIPTION</td>";
            displayoutput=displayoutput+"<td align='center' bgcolor='#999999'>BALANCE</td>";
            decimal sum=0.0M;

		string strConnect = "server=AUCKLAND;uid=gl1197;pwd=FDT24tccG;database=gl1197";

		// specify the SELECT statement to extract the data

		string strSelect = "SELECT * FROM glmaster";
		outSelect.InnerText = strSelect;		// and display it

		try
		{
			// create a new Connection object using the connection string
			SqlConnection objConnect = new SqlConnection(strConnect);

			// open the connection to the database
			objConnect.Open();

			// create a new Command using the connection object and select statement
			SqlCommand objCommand = new SqlCommand(strSelect, objConnect);

			// declare a variable to hold a DataReader object
			SqlDataReader objDataReader;

			// execute the SQL statement against the command to fill the DataReader
			objDataReader = objCommand.ExecuteReader();

			// iterate through the records in the DataReader getting field values
			// the Read method returns False when there are no more records
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

			// close the DataReader and Connection
			objDataReader.Close();
			objConnect.Close();
		}
		catch (Exception objError)
		{
			// display error details
			outError.InnerHtml = "<b>* Error while accessing data</b>.<br />"
							+ objError.Message + "<br />" + objError.Source;
			return;		//  and stop execution
		}

		// add closing table tag and display the results
		displayoutput += "</table><p>Sum of Balances=";
		displayoutput += sum;
		displayoutput += "<p>Row Count="+c.ToString();
		DateTime t = DateTime.Now;
                displayoutput += "<P>Normal Termination "+t;

		outResult.InnerHtml = displayoutput;

	}
	
</script>
</body>
</html>



