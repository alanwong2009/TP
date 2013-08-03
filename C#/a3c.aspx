<%@Page Language="C#"%>
<%@Import Namespace="System.Data" %>
<%@Import Namespace="System.Data.SqlClient" %>
<html>
<body>
<p>Alan Wong
<p>0847623
<p>gl1197
<div> SELECT command: <b><span id="outSelect" runat="server"></span></b></div>
<div id="select1" runat="server"> </div>
<div>SRN command: <b><span id="outsrn" runat="server"></span></b></div>
<div id="srn" runat="server"> </div>
<div>Update command: <b><span id="outupdate" runat="server"></span></b></div>
<div id="update" runat="server"></div>
<div>Insert command: <b><span id="outinsert" runat="server"></span></b></div>
<div id="insert" runat="server"></div>
<p><input type="button" onClick="parent.location='index.htm'" value="Return to Main Menu">

<script language="C#" runat="server">
    void Page_Load(Object sender, EventArgs e)
    {
        Response.Write("<p>Begin...");
        int nr = 0;
        int rc = 0;     
        int ec = 0;
        String Sql;
        String login = "server=AUCKLAND;uid=gl1197;pwd=FDT24tccG;database=gl1197";

        try
        {
            

            for (int i = 0; i < 6; i++)
            {
                String poop = Request.Form["a00" + i + "1field"];
                if (poop != "")
                {
                    nr = nr + 1;
                }
            }

          
            
            for (int j = 0; j < nr ; j++)
            {
                SqlConnection objConnect = new SqlConnection(login);
                objConnect.Open();
            
                Response.Write("<p>Connection Opened");
                
                Sql = "select * from glmaster where ";
                Sql = Sql + "major=" + Request.Form["a00" + j + "1field"] + " and ";
                Sql = Sql + "minor=" + Request.Form["a00" + j + "2field"] + " and ";
                Sql = Sql + "sub1=" + Request.Form["a00" + j + "3field"] + " and ";
                Sql = Sql + "sub2=" + Request.Form["a00" + j + "4field"];
                outSelect.InnerHtml = Sql;
                
                SqlCommand SqlCommand = new SqlCommand(Sql, objConnect);
                SqlDataReader SqlDataReader = SqlCommand.ExecuteReader();

                while (SqlDataReader.Read())
                {
                    rc = rc + 1;
                }
               
                if (rc == 0)
                {
                    ec = ec + 1;
                    select1.InnerHtml = "<p>One Account Not Found for "+Sql;
                }
                SqlDataReader.Close();
                objConnect.Close();
                rc = 0;
            }
            
        }
        catch (Exception select1Error)
        {
            ec = ec + 1;
            select1.InnerHtml = "<b>* Error </b>.<br />" + select1Error.Message + "<br />" + select1Error.Source;
        }
        
        

        if (ec == 0)
        {
            Response.Write("<p>Finished Checking the Account Existence");
            Response.Write("<p>Begin checking SRN...");
            String sqlsrn = "select * from je where sourceref=" + Request.Form["sourceref"];
            outsrn.InnerHtml = sqlsrn;
            try
            {
                SqlConnection newConnect = new SqlConnection(login);
                newConnect.Open();
                Response.Write("<p>Connection Opened");
                SqlCommand srncommand = new SqlCommand(sqlsrn, newConnect);
                SqlDataReader srndatareader = srncommand.ExecuteReader();
                
                while (srndatareader.Read())                    
                {
                    rc = rc + 1;
                }
                
                srndatareader.Close();
                newConnect.Close();
            }
            catch (Exception srnerror)
            {
                ec = ec + 1;
                srn.InnerHtml = "<b>* Error </b>.<br />" + srnerror.Message + "<br />" + srnerror.Source;
            }
        }

        if (rc != 0)
        {
            ec = ec + 1;
            srn.InnerHtml = "<p>SRN already exist<p>";
        }
      
        if (ec == 0)
        {
            Response.Write("<p>Finished checking SRN existence");
            Response.Write("<p>Begin Update the balance");
            String sqlupdate;
            
            for (int i = 0; i < nr; i++)
            {
                sqlupdate = "UPDATE glmaster SET Balance=balance+";
                sqlupdate = sqlupdate + Request.Form["a00"+i+"5field"] + " WHERE ";
                sqlupdate = sqlupdate + "major=" + Request.Form["a00"+i+"1field"] + " and ";
                sqlupdate = sqlupdate + "minor=" + Request.Form["a00"+i+"2field"] + " and ";
                sqlupdate = sqlupdate + "sub1=" + Request.Form["a00"+i+"3field"] + " and ";
                sqlupdate = sqlupdate + "sub2=" + Request.Form["a00"+i+"4field"];

                outupdate.InnerHtml = sqlupdate;

                String ad = Request.Form["com0"+i];
                if (ad == "")
                {
                    ad = "No Description";
                }
                String sqlinsert = "Insert into je values(";
                sqlinsert = sqlinsert + Request.Form["sourceref"] + ", ";
                sqlinsert = sqlinsert + (i + 1) + ", ";
                sqlinsert = sqlinsert + Request.Form["a00" + i + "1field"] + ", ";
                sqlinsert = sqlinsert + Request.Form["a00" + i + "2field"] + ", ";
                sqlinsert = sqlinsert + Request.Form["a00" + i + "3field"] + ", ";
                sqlinsert = sqlinsert + Request.Form["a00" + i + "4field"] + ", ";
                sqlinsert = sqlinsert + "'" + ad + "'" + ", ";
                sqlinsert = sqlinsert + Request.Form["a00" + i + "5field"] + ")";

                outinsert.InnerHtml = sqlinsert + "<br>";

                try
                {
                    SqlConnection updateConnect = new SqlConnection(login);
                    SqlCommand updatecommand = new SqlCommand(sqlupdate, updateConnect);
                    updateConnect.Open();
                    updatecommand.ExecuteNonQuery();
                    updateConnect.Close();
                }
                catch (Exception updateerror)
                {
                    ec = ec + 1;
                    update.InnerHtml = "<b>* Error </b>.<br />" + updateerror.Message + "<br />" + updateerror.Source;
                }
                Response.Write("<p>Finished Updating");
                try
                {
                    SqlConnection insConnect = new SqlConnection(login);
                    SqlCommand insertcommand = new SqlCommand(sqlinsert, insConnect);
                    insConnect.Open();
                    insertcommand.ExecuteNonQuery();
                    insConnect.Close();
                }
                catch (Exception inserterror)
                {
                    ec = ec + 1;
                    insert.InnerHtml = "<b>* Error </b>.<br />" + inserterror.Message + "<br />" + inserterror.Source;
                }
                Response.Write("<p>Insert into JE successed");
            }

        }
    }
</script>
</body>
</html>
