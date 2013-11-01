<%@ Language=VBScript %>
<%
        If isEmpty(Session("StudentID")) then
                Response.Redirect "./default.asp" 
        End If
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 TRANSITIONAL//EN">
<html>
<head>
<title>Lake Forest MBA Writing Assessment: Grammar 1</title>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<META NAME="rating" CONTENT="General">
<META NAME="revisit-after" CONTENT="15 days">
<META HTTP-EQUIV = "expires" CONTENT="Never">
<META NAME="ROBOTS" CONTENT="ALL">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META name="description" content="Lake Forest MBA Writing Assessment: Grammar 1">
<META name="keywords" Content="writing at work effective business writing grammar puncutation">
<link rel="stylesheet" type="text/css" href="master.css">
</head>

<body>
 
<!-- This table sets the format for the page -->
  <table cellpadding="0" cellspacing="0" border="0" width="600" cols="3">
   <tr valign="top" align="left">
   
   <!-- This column displays the logos -->
        <td  class="small" width="100" align="center"><img src="./assets/logo_lfmba1.gif" width="77" height="90" border="0"><br><br>
	   <!-- # include file ="leftmenu.inc"--.
        </td>


   <!-- This column inserts a spacer -->        
        <td width="10">
        <img src="./assets/spacer.gif" width="10" border="0"></td>
        
   <!-- This column contains the headings and text -->  
		<td width="490"><img src="./assets/logo_lfmba2.gif" width="162" height="90" border="0">
		<h1>Writing Skills Assessment</h1>
                        
                        <!-- This contains the form for the fifth question-->
                        <form method="Post" action="Gq1_submit.asp">
                        <h2>Mechanics: Grammar and Punctuation &#151; Grammar and Sentence Structure</h2>
						<h3>Question 1</h3>
                        <p>This sample appears in all six questions in this grammar section. In each one, a different sentence is highlighted.
                        Please read the highlighted sentence and select the best answer in the <b>Revision box</b> that 
                        follows the memo.</p>
                        <img src="./assets/sample5A.gif" width="485" height="465" border="0">
                        <!-- This contains the Assessment for the Sample -->
                        <h2>Sentence 1 Revision</h2>
                        <p>Select 
                        either the appropriate correction below or mark it "OK as is."</p>
                        <p>
                                                
                       <table cellpadding="2" cellspacing="0" border="1" width="480" cols="3">
                        <tr>
                        <td align="center" width="35"><b>(A)</b></td>
                        <td align="center" width="35">
                        <input type="radio" name="Gq1" value="A"
                        <% if Session("Gq1") = "A" then %> CHECKED <% end if %>></td>
                        <td class=head align="left" width="390">
                        <b>We are updating our Marketing materials, and could use your help 
                        to provide the most recent experience of our professional staff.</b></td>
                        </tr>                   
                        <tr>
                        <td align="center" width="35"><b>(B)</b></td>
                        <td align="center" width="35">
                        <input type="radio" name="Gq1" value="B"
                        <% if Session("Gq1") = "B" then %> CHECKED <% end if %>></td>
                        <td class=head align="left" width="390">
                        <b>We are updating our Marketing materials; and could use your help 
                        to provide the most recent experience of our professional staff.</b></td>
                        </tr>                   
                        <tr>
                        <td align="center" width="35"><b>(C)</b></td>
                        <td align="center" width="35">
                        <input type="radio" name="Gq1" value="C"
                        <% if Session("Gq1") = "C" then %> CHECKED <% end if %>></td>
                        <td class=head align="left" width="390">
                        <b>We are updating our Marketing materials and could use your help, 
                        to provide the most recent experience of our professional staff.</b></td>
                        </tr>                                   
                        <tr>
                        <tr>
                        <td align="center" width="35"><b>(D)</b></td>
                        <td align="center" width="35">
                        <input type="radio" name="Gq1" value="D"
                        <% if Session("Gq1") = "D" then %> CHECKED <% end if %>></td>
                        <td class=head align="left" width="390">
                        <b>OK as is</b></td>
                        </tr>                                   
                        </table>
<!-- This contains area for comments or recommendations 
                        <h2>Comments or Questions?</h2>  
                        <textarea name="Comments5A" rows="5" cols="35" wrap><% =Session("Comments5A") %></textarea>                             
 This closes out the row  -->


        </td>
        </tr>



   <!-- This column displays the Writing at Work logo -->
        <tr><td width="100" align="center">&nbsp;
        </td>

   <!-- This column inserts a spacer -->        
        <td width="10">
        <img src="./assets/spacer.gif" width="10" border="0"></td>
        
   <!-- This submits the form information -->   
        <td width="490" valign="middle" align="left"><p><input type="submit" value="Next" name="Action">
                                <input type="reset" value="Clear">
                                </p><p>Assessment developed by Writing at Work.</p></td>
        </tr>                   
                                </p>
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
