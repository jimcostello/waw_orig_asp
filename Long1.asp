<%@ Language=VBScript %>
<%
        If isEmpty(Session("StudentID")) then
                Response.Redirect "./default.asp" 
        End If
		
		dim strPagename
		strPagename="Long1"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 TRANSITIONAL//EN">
<html>
<head>
<title>Lake Forest MBA Writing Assessment: Outline</title>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<META NAME="rating" CONTENT="General">
<META NAME="revisit-after" CONTENT="15 days">
<META HTTP-EQUIV = "expires" CONTENT="Never">
<META NAME="ROBOTS" CONTENT="ALL">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META name="description" content="Lake Forest MBA Writing Assessment: Outline">
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
						<h3>Construct an Outline</h3>
                        <p>An outline provides a formal organizing structure for your writing.
						The elements of an outline&#151;purpose, main ideas, supporting points,
						conclusion&#151;can be organized in several ways:</p>
						<ul>
						<li><strong>Topical.</strong> Begin with your purpose. Then list  
						each of the main ideas in sequence with their supporting points 
						indented underneath them. End with your conclusion. This outline 
						applies particularly well to situations where you are attempting to 
						state a case or persuade your audience.</li>
						<li><strong>Chronological.</strong> Your topic may naturally lend itself 
						to this type of organization. If you are reciting a history of events 
						(such as in  Project Report) where b follows a and c follows b, then this 
						outline makes the most sense.</li>
						<li><strong>Process.</strong> This is similar to a chronological outline, 
						but with steps in a process as the organizing principle. It's often a 
						"cookbook" approach: step 1, step 2, step 3, and so on.  This works 
						well when describing any process or procedure such as a sales process, 
						reimbursement procedure, and so on.</li>
						<li><strong>Compare and Contrast.</strong>Here you present one thing 
						against another to promote the advantages of one. This can be effective 
						when you are writing a proposal or recommendation . It is typically used 
						within one of the other organizing principles such as topical or process.</li>
						<li><strong>Cause and Effect.</strong> This can be very effective when you 
						are trying to explain something or convince someone. For example, if you 
						are writing a report to analyze sales trends, you can build cause and effect 
						relationships to make your points. As with compare and contrast, it is used 
						within other organizing principles. </li>
						</ul>
						<p>In this exercise, you will construct an outline for a report based on the 4-page meeting 
						minutes provided. The minutes, themselves, seem to be a bit skimpy, so all the 
						more important that the outline emphasizes the important information. The process 
						is a bit different for these longer exercises, so follow the steps outlined below:</p> 
						<ol>
						<li><a href="MeetingMinutes.pdf">Download the minutes (PDF)</a> and review them on your own.</li>
						<li>Construct an outline for the report based on the minutes. Do this offline from this 
						web site in your own word processing software, such as Microsoft Word.&#153; (Note that if you are 
						away from this web site for more than 20 minutes, you will need to sign in again.) </li>
						<li>When you finish your outline, <a href="Long1A.asp">compare it to the model provided
						on the next page</a> and rate yourself.
						</ol>
						
						<div class="samplebox3">
						<strong>Background</strong>
						<p>You work in the sales department that supplies a regional chain of grocery 
						stores.  This store has declared bankruptcy and has invited you to attend a 
						meeting that will divulge the company's status.</p>
						<p>Representing one of the vendors invited, you created the minutes to 
						distribute to your regional sales team.  Now you need to write a report based
						on these minutes. Your first step will be to organize the information in the
						minutes by constructing an outline for the report.   </p> 
						</div>
						<p><a href="./Long1A.asp">Go to Outline Model and compare it to your own.</a>  

                             
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
