<%@ Language=VBScript %>
<%
        If isEmpty(Session("StudentID")) then
                Response.Redirect "./default.asp" 
        End If
		
		dim strPagename
		strPagename="Long3A"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 TRANSITIONAL//EN">
<html>
<head>
<title>Lake Forest MBA Writing Assessment: Email</title>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<META NAME="rating" CONTENT="General">
<META NAME="revisit-after" CONTENT="15 days">
<META HTTP-EQUIV = "expires" CONTENT="Never">
<META NAME="ROBOTS" CONTENT="ALL">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META name="description" content="Lake Forest MBA Writing Assessment: Email">
<META name="keywords" Content="writing at work effective business writing grammar puncutation">
<link rel="stylesheet" type="text/css" href="master.css">
</head>

<body>
<div id="wrap">
 
<!-- This table sets the format for the page -->
  <table cellpadding="0" cellspacing="0" border="0" width="720" cols="5">
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
                        <h2>Effective Writing Principles</h2>
						<h3>Email Model</h3>
                        <p>Shown below is a model for an email.</a></p>
						
						<div class="samplebox3">
						For your assessment, consider these questions:
						<ul>
						<li>Did you use the hourglass or pyramid model of organization?</li>
						<br/><br/>
						<li>Does your opening paragraph start with a negative?</li> 
						<br/><br/>
						<li>Is the message free of clichés and deadwood?</li>
						<br/><br/>
						<li>Does the messag subordinate negatives, highlight positives&#151;stress 
						what readers <em>can do</em> rather than <em>cannot</em>?</li>
						<br/><br/>
						<li>Does the message provide details and reasoned argument rather than 
						appeals to authority or unsupported generalizations? </li>
						<br/><br/>
						<li>How strongly would this message have convinced you? </li>
						
						</ul>

						</div>
						
						<h2>Rate Yourself</h2>
						<p>Use the chart below to rate your email. Select Level 1, Level 2, Level 3, or Level 4 and click <strong>Submit</strong>.
						You may <a href="Long3AComponent.asp">view a table explaining the different levels</a> before submitting your rating.</p>
						
						<form action="Long3Rate.asp" method="post">
						<table cellpadding="2" cellspacing="0" border="1" width="490" cols="5">
                        <tr bgcolor="FFFFE0">
                        <td align="center" valign="top" width="75"><b>Component</b></td>
                        <td align="center" valign="top" width="100"><b>Level 1 <br />  Minimal</b></td>
                        <td align="center" valign="top" width="100"><b>Level 2 <br /> Emerging</b></td>
                         <td align="center" valign="top" width="100"><b>Level 3 <br /> Competent</b></td>
                         <td align="center" valign="top" width="100"><b>Level 4 <br /> Exemplary</b></td>
                        </tr>  

						<tr bgcolor="FFFFE0">
                        <td align="center" valign="top" width="75"><b>Your Rating</b></td>
                        <td align="center" valign="top" width="100"><input type="radio" name="Long3Rate" value="1"
                        <% if Session("Long3Rate") = "1" then %> CHECKED <% end if %></td>
                        <td align="center" valign="top" width="100"><input type="radio" name="Long3Rate" value="2"
                        <% if Session("Long3Rate") = "2" then %> CHECKED <% end if %></td>
                         <td align="center" valign="top" width="100"><input type="radio" name="Long3Rate" value="3"
                        <% if Session("Long3Rate") = "3" then %> CHECKED <% end if %></td>
                         <td align="center" valign="top" width="100"><input type="radio" name="Long3Rate" value="4"
                        <% if Session("Long3Rate") = "4" then %> CHECKED <% end if %></td>
                        </tr>
                                   
                        </table>
						<p><input type="submit" value="Next" name="Action">
                                <input type="reset" value="Clear"></p>
						</form>
						
						 <p>&nbsp;</p>
						
						
 
						  

                             
<!-- This closes out the row  -->
        </td>
				
		
   <!-- This column inserts a spacer -->        
        <td width="10">
        <img src="./assets/spacer.gif" width="10" border="0"></td>
		
   <!-- This column displays the logos -->
        <td  class="small" width="100" align="center">
		<img src="./assets/spacer.gif" height="90"  width="100" border="0">
		<br><br>
	   <!-- #include file ="rightmenu.inc"-->
        </td>	
        </tr>
		
<!-- #include file ="bottom.inc"-->

<!-- This closes out the table -->
        </table>

</div>


</body>
</html>

