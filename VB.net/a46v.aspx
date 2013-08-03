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
<b>The Simple Table Listing Program Using a Data Grid<br>in VB.NET<hr>Lists the Contents of <i>JE</i><br>Using a DataReader in SQL
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
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>SOURCEREF</td>"
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>SRSEQ</td>"
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>JEMAJOR</td>"
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>JEMINOR</td>"
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>JESUB1</td>"
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>JESUB2</td>"
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>JEDESC</td>"
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>JEAMOUNT</td>"
    
            Dim	 strConnect as String = "server=AUCKLAND;uid=gl1197;pwd=FDT24tccG;database=gl1197"

        Dim strSelect As String = "SELECT * FROM je WHERE SOURCEREF = " & SRN.Text
        
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
                displayoutput += "<tr><td>" + objDataReader("SOURCEREF").ToString + "</td><td>" _
                + objDataReader("SRSEQ").ToString + "</td><td>" _
                + objDataReader("JEMAJOR").ToString + "</td><td>" _
                + objDataReader("JEMINOR").ToString + "</td><td>" _
                + objDataReader("JESUB1").ToString + "</td><td>" _
                + objDataReader("JESUB2").ToString + "</td><td>" _
                + objDataReader("JEDESC").ToString + "</td><td align='right'>" _
                + objDataReader("JEAMOUNT").ToString + "</td></tr>"
             Loop

'			// close the DataReader and Connection
			objDataReader.Close()
			objConnect.Close()
		
		Catch objError as Exception 
		
'			// display error details
			outError.InnerHtml = "<b>* Error </b>.<br />"+ objError.Message + "<br />" + objError.Source

		end Try

'		// add closing table tag and display the results		
		displayoutput += "<p>Row Count="+c.ToString
                displayoutput += "<P>Normal Termination "+Now.toString


		outResult.InnerHtml = displayoutput

end sub	
	
</script>
<form id="Form1" method="post" runat="server">
<p>SRN <asp:TextBox columns="4" ID="SRN" Runat="server" />
<asp:RequiredFieldValidator id="RequiredFieldValidator1" runat="server" EnableClientScript="False" Width="208px" Display="Dynamic" ControlToValidate="SRN" ErrorMessage="SRN is Required"></asp:RequiredFieldValidator>  
<asp:RangeValidator id="RangeValidator1" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="SRN" ErrorMessage="SRN MUST BE BETWEEN 0-9999" MinimumValue="0" MaximumValue="9999" Type="Integer"></asp:RangeValidator>   
<asp:CompareValidator id="CompareValidator1" runat="server" EnableClientScript="False" Display="Dynamic" ControlToValidate="SRN" ErrorMessage="Must be digits 0-9 only" Operator="DataTypeCheck" Type="Integer"></asp:CompareValidator> <br />
<p><asp:Button Text="Query Account" ID="btnquery" OnClick="Page_Load" Runat="server" />
<p><asp:Label ID="numa" ForeColor="#ff0000" Runat=Server/>
</form>
</body>
</html>



