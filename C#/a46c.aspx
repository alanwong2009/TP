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
<b>The Simple Table Listing Program Using a Data Grid<br>in VB.NET<hr>Lists the Contents of <i>JE</i><br>Using a DataReader in SQL
</td>
</tr></table>

<div>SELECT command: <b><span id="outSelect" runat="server"></span></b></div>
<div id="outError" runat="server"> </div>
<div id="outResult" runat="server"></div>

<script language="C#" runat="server">

    void Page_Load(Object sender, EventArgs e)
    {
        int c = 0;
        String displayoutput = "<center><table border='1'>";
        displayoutput = displayoutput + "<tr>";
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>SOURCEREF</td>";
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>SRSEQ</td>";
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>JEMAJOR</td>";
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>JEMINOR</td>";
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>JESUB1</td>";
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>JESUB2</td>";
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>JEDESC</td>";
        displayoutput = displayoutput + "<td align='center' bgcolor='#999999'>JEAMOUNT</td>";

        String strConnect = "server=AUCKLAND;uid=gl1197;pwd=FDT24tccG;database=gl1197";

        String strSelect = "SELECT * FROM je WHERE SOURCEREF = " + SRN.Text;

        outSelect.InnerText = strSelect;
        try
        {
            SqlConnection objConnect = new SqlConnection(strConnect);
            objConnect.Open();
            SqlCommand objCommand = new SqlCommand(strSelect, objConnect);
            SqlDataReader objDataReader = objCommand.ExecuteReader();
            while (objDataReader.Read())
            {
                c = c + 1;
                displayoutput += "<tr><td>" + objDataReader["SOURCEREF"] + "</td><td>" + objDataReader["SRSEQ"] + "</td><td>" + objDataReader["JEMAJOR"] + "</td><td>" + objDataReader["JEMINOR"] + "</td><td>" + objDataReader["JESUB1"] + "</td><td>" + objDataReader["JESUB2"] + "</td><td>" + objDataReader["JEDESC"] + "</td><td align='right'>" + objDataReader["JEAMOUNT"] + "</td></tr>";
            }
            objDataReader.Close();
            objConnect.Close();
        }
        catch (Exception objError)
        {
            outError.InnerHtml = "<b>* Error </b>.<br />" + objError.Message + "<br />" + objError.Source;
        }

        displayoutput += "<p>Row Count=" + c.ToString();
        DateTime t = DateTime.Now;
        displayoutput += "<P>Normal Termination " + t;

        outResult.InnerHtml = displayoutput;

    }
	
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



