<%@ Language=VBScript %>
<% Option Explicit %>
<%
dim conn, TheMessage, RSStudent
if isempty(Request.Form("EmailAddress")) then
	'first time entering page
	TheMessage = "If you have previously registered,  enter your email address and password to begin. <br><br>If you are not registered, click on the New Student graphic below.</p>"
elseif Request.Form("EmailAddress") = "reports" then
		Response.Redirect "./reports.asp"

else
	set conn = server.createobject ("adodb.connection")
	conn.open "writingatworkonline"
	set RSStudent = conn.execute ("SELECT StudentID, StudentFirstName, StudentLastName, StudentEmailAddress FROM students where StudentEmailAddress = '"_
		 & Request.Form("EmailAddress") & "' and StudentPassword = '" _
		 & Request.Form("Password") & "'")

	if RSStudent.EOF then
		'invalid login
		TheMessage = "Your email address and password are not registered in our database. Try entering your email name and password again or  click on the New Student graphic below to register."	
	'-- This closes the recordset and database connection
		RSStudent.Close
		conn.Close
		Set conn = Nothing	
	else
		'valid entry
		Session("StudentID") = RSStudent("StudentID")
		Session("FirstName") = RSStudent("StudentFirstName")
		Session("LastName") = RSStudent("StudentLastName")
		Session("EmailAddress") = RSStudent("StudentEmailAddress")
		'-- This closes the recordset and database connection
		RSStudent.Close
		conn.Close
		Set conn = Nothing	
		Response.Redirect "./begin.asp"
	end if
end if		
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 TRANSITIONAL//EN">
<html>
<head>
<title> Advanced Business Writing: Login</title>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<META NAME="rating" CONTENT="General">
<META NAME="revisit-after" CONTENT="15 days">
<META HTTP-EQUIV = "expires" CONTENT="Never">
<META NAME="ROBOTS" CONTENT="ALL">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META name="description" content="Unilever Effective Business Writing Seminar: Login">
<META name="keywords" Content="writing at work effective business writing grammar puncutation">
<link rel="stylesheet" type="text/css" href="unilever.css">
</head>

<body>

 
<!-- This table sets the format for the page -->
  <table cellpadding="0" cellspacing="0" border="0" width="600" cols="3">
  <form action="./default.asp"  method=post>
   <tr valign="top" align="left">
   
   <!-- This column displays the logos -->
	<td width="100" align="center"><img src="./assets/logo_reserve.jpg" width="99" height="92" border="0"></p><p>
	</td>

   <!-- This column inserts a spacer -->	
	<td width="10">
	<img src="./assets/spacer.gif" width="10" border="0"></td>
	
   <!-- This column contains the headings and text -->	
	<td width="490"><img src="./assets/logo_reserve2.gif" width="249" height="54" border="0">
	<h2>Advanced Business Writing <h2>
  
	<!-- This part prints the appropriate message -->
	<p class="bold">
	<% Response.Write TheMessage %>
	</p>
  
	<!-- This embedded table sets the form text boxes and the new student icon and the submit button-->
	<table cellpadding="0" cellspacing="0" border="0" width="490" cols="2">
	<tr valign="top" align="left">
	<td width="390">
	<p>Email Address<br>
	<input type="text" name="EmailAddress" value="" size="30" maxlength="50"></p>	
	<p>Password<br>
	<input type="password" name="Password" value="" size="20" maxlength="50"></p></td>
	<td width="100" align="center"><a href="./register.asp"><img src="./assets/pencil2.gif" heigh="50" width="50" border="0" alt="New Student"></a><br>
	<b>New Student</b><!-- This closes out the row  -->
	</table>

	</td>
	</tr>

   <!-- This column displays the Writing at Work logo -->
	<tr><td width="100" align="center" valign="top">
	<img src="./assets/logo_waw.gif" width="78" height="64" border="0"></p></td>

   <!-- This column inserts a spacer -->	
	<td width="10">
	<img src="./assets/spacer.gif" width="10" border="0"></td>
	
   <!-- This submits the form information -->	
	<td width="490" valign="top" align="left"><p><input type="submit" value="Next" name="Action">
				<input type="reset" value="Clear">
				<p>Course developed by Writing at Work.</p></td>
	</tr>			
			</form>
			


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
