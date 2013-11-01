<%@ Language=VBScript %>
<%
	If isEmpty(Session("StudentID")) then
		Response.Redirect "./default.asp" 
	End If
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 TRANSITIONAL//EN">
<html>
<head>
<title>Advanced Business Writing: Welcome</title>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<META NAME="rating" CONTENT="General">
<META NAME="revisit-after" CONTENT="15 days">
<META HTTP-EQUIV = "expires" CONTENT="Never">
<META NAME="ROBOTS" CONTENT="ALL">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META name="description" content="Unilever Effective Business Writing seminar: Welcome">
<META name="keywords" Content="writing at work effective business writing grammar puncutation">
<link rel="stylesheet" type="text/css" href="unilever.css">
</head>

<body>


<!-- This table sets the format for the page -->
  <table cellpadding="0" cellspacing="0" border="0" width="600" cols="3">
   <tr valign="top" align="left">
   
   <!-- This column displays the logos -->
	<td width="100" align="center">
	<img src="./assets/logo_reserve.jpg" width="99" height="92" border="0"><br>
	<hr width="90%">
	<b>JUMP TO:</b><p>
	<a href="./begin.asp">Welcome</a></p><p>
	<a href="./question1.asp">Sample 1</a><br>
	<a href="./question2.asp">Sample 2</a><br>
	<a href="./question3.asp">Sample 3</a><br>
	<a href="./question4.asp">Sample 4</a><br>
	<a href="./results.asp">Results</a></p>
	<hr>
	<a href="./question1.asp"><img src="./assets/begin.jpg" width="100" height="40" border="0"></a>
	</td>

   <!-- This column inserts a spacer -->	
	<td width="10">
	<img src="./assets/spacer.gif" width="10" border="0"></td>
	
   <!-- This column contains the headings and text -->	
	<td width="490">
	<img src="./assets/logo_reserve2.gif" width="249" height="54" border="0">
	<h1><img src="./assets/welcome.gif" width="124" height="36" align="bottom"></h1>	
	<img src="./assets/penpaper3.gif" width ="65" height="73" align="right">
	<p>To get a head start on the follow-up workshop for <i>Advanced Business Writing</i> this January 23, this online session contains several short writing excerpts to review and a questionnaire. Please take about 15 minutes to review the samples and complete the questionnaire. For writing samples 1-4, you will be asked to evaluate effectiveness and suggest changes. </p>
	<p>Be sure to click the Next button	at the bottom of each sample page to record your work. Also, be sure to continue through all questions without pausing for an extended period (10 minutes or more).
	This inactivity will automatically log you off the session as a security precaution. Your responses
	will not be recorded, and you will need to sign on again.</p><p>
	<span class="bold">Your work will be emailed directly to the instructor, Steven Schultz.  As in all of your work submitted or completed for the seminar, your responses will be held confidential.</span></p>
	<p>Please print copies of the samples you edit here and bring them to the session. You may print these pages 
   as you go along or print the summary page that appears when you complete the entire exercise.</p>
	<p> The "Jump To" menu at the left lets you go directly to any screen within this section to review or correct your work. Click the <a href="./question1.asp">Begin</a> button at the left to begin your session.</p> 

<table cellpadding="0" cellspacing="0" border="0" width="490" cols="4">
<tr>

<td width="100" rowspan="3" align="left" ><img src="./assets/logo_waw.gif" width="78" height="64" border="0"><br></td>
<td width="130" align="left"></td>
<td width="130" align="left"></td>
<td width="130" align="left"> </td>
</tr>
<tr>
<td width="130" align="left">Tampa, Florida</td>
<td width="130" align="left">(813) 854-1999</td> 
<td width="130" align="left">(Fax) 818-8134</td>
</tr>
<tr>
<td width="130" align="left">E-mail:</td>
<td align="left" colspan="2"> <a href="mailto:writing@tampabay.rr.com">writing@tampabay.rr.com
</a></td>
</tr>
</table>

  


<!-- This closes out the row  -->

	</td>
	</tr>

	<!-- This draws a horizontal line -->
	<tr><td colspan="3"><hr></td></tr>

   <!-- This column displays the ActionWeb Design logo -->
	<td width="100" align="center">
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
