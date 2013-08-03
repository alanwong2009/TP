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
        Dim strConnect As String = "server=AUCKLAND;uid=gl1197;pwd=FDT24tccG;database=gl1197"

            Dim strSelect as String = "SELECT * FROM glmaster"
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
</body>
</html>



