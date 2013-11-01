<%@ Language=VBScript %>
<%
        If isEmpty(Session("StudentID")) then
                Response.Redirect "./default.asp" 
        End If
'-- this stores the info from the first question

dim strComments2,conn,RSTestsSelect						


set conn = server.createobject ("adodb.connection")
                        conn.open "writingatworkonline" 
						
						
					set RSTestsSelect = conn.execute ("SELECT Comments2 FROM lfgsm_tests where StudentID=" & Session("StudentID") & "")


                '-- This closes the recordset and database connection
                conn.Close
                Set conn = Nothing 
				
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 TRANSITIONAL//EN">
<html>
<head>
<title>Lake Forest MBA Writing Assessment: Writing 2</title>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<META NAME="rating" CONTENT="General">
<META NAME="revisit-after" CONTENT="15 days">
<META HTTP-EQUIV = "expires" CONTENT="Never">
<META NAME="ROBOTS" CONTENT="ALL">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META name="description" content="Lake Forest MBA Writing Assessment: Writing 2">
<META name="keywords" Content="writing at work effective business writing grammar puncutation">
<link rel="stylesheet" type="text/css" href="master.css">
</head>

<body>
 
<!-- This table sets the format for the page -->
  <table cellpadding="0" cellspacing="0" border="0" width="600" cols="3">
   <tr valign="top" align="left">
   
 <!-- This column displays the logos -->
        <td class="small" width="100" align="center"><img src="./assets/logo_lfmba1.gif" width="77" height="90" border="0"><br><br>
			 <!-- #include file ="leftmenu.inc"-->
        </td>

   <!-- This column inserts a spacer -->        
        <td width="10">
        <img src="./assets/spacer.gif" width="10" border="0"></td>
        
   <!-- This column contains the headings and text -->  
		<td width="490"><img src="./assets/logo_lfmba2.gif" width="162" height="90" border="0">
		<h1>Writing Skills Assessment</h1>
                        
                        <!-- This contains the form for the fifth question-->
                        <form method="Post" action="Short2A.asp">
                       <h2>The Writing Process</h2>
						<h3>Sample 2</h3>
                        <p>Review the passage below and rewrite it in the empty text box to eliminate the wordiness.   In your rewrite, do not eliminate essential content or change intended meaning.  </p>
                        <div class="samplebox">
                        This email will confirm your enrollment in our Motor Vehicle Dealer Class scheduled for December 7.  Don Larson, the Carol County Tax Collector, offers this class free of charge as a courtesy to the personnel who work for motor vehicle dealers and for financial institutions in Carol and surrounding counties.  
<p>
I have attached a copy of the outline which shows the forms, rules, procedures and other information to be covered in the class and a map showing the general vicinity of our Mid-County Service Center.  The time frame for this class is from 9:00 a.m. to 4:00 p.m.  Please note that there will be a short break in the morning and in the afternoon, as well as a one-hour break for lunch.</p>
<p>
Please note that this class is not the required continuing education class which the State requests dealerships to take in order to renew their license.  This particular class covers the forms which you process in the Tax Collector’s dealer sections, as well as the State’s rules and procedures which apply to these forms.</p>
<p>
I would appreciate receiving a reply e-mail to acknowledge your receipt of this communication and the attachments.</p>
<p>
Again, thank you for your interest and please feel free to call me at (777) 777-7777 if you have any questions.</p>
                        </div>
                        
                        
<!-- This contains area for comments or recommendations -->
                        <h2>Rewrite Your Version Below:</h2>  
<textarea name="Comments2" rows="5" cols="55" wrap="hard"><% =Replace(RSTestsSelect("Comments2"),"*","'") %></textarea><!-- This closes out the row  -->

							<p><input type="submit" value="Next" name="Action">
                                <input type="reset" value="Clear"></p>

						                                
                        
						</form>

        </td>
        </tr>
 <!-- This inserts a blank line -->
        <tr><td colspan="3">&nbsp;</td></tr>


   <!-- This column displays the Writing at Work logo -->
        <tr><td width="100" align="center">&nbsp;
        </td>

   <!-- This column inserts a spacer -->        
        <td width="10">
        <img src="./assets/spacer.gif" width="10" border="0"></td>
        
   <!-- This submits the form information -->   
        <td width="490" valign="middle" align="left">
                                </p><p>Assessment developed by Writing at Work.</td>
        </tr>                   
                      


        <!-- This draws a horizontal line -->
        <tr><td colspan="3"><hr></td></tr>

    <!-- This column displays the ActionWeb Design logo -->
        <tr><td width="100" align="center">
        &nbsp;</p></td>

   <!-- This column inserts a spacer -->        
        <td width="10">
        <img src="./assets/spacer.gif" width="10" border="0"></td>
        
   <!-- This column contains the headings and text -->  
        <td width="490" valign="top" align="left">Web site designed by Writing at Work, Web Design and Interactive Training Group.</td>
        </tr>

<!-- This closes out the table -->
        </table>


</body>
</html>
