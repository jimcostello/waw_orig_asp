<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file="adovbs.inc" -->
<%

'-- This declares variables, defines msg query string, saves session variables, and opens datbase connection
Dim msg, Query, Open, conn, RSStudent,RSTestsSelect,RSTestsInsert, RSPosition
msg = Request.QueryString ("msg")

Session("FirstName") = Request.Form("FirstName")
Session("LastName") = Request.Form("LastName")
Session("EmailAddress") = Request.Form("EmailAddress")
Session("Password") = Request.Form("Password")
Session("RepeatPassword") = Request.Form("RepeatPassword")

set conn = server.createobject ("adodb.connection")
conn.open "writingatworkonline"
                
'-- this checks to see status of form entry and submission
if isempty(Request.Form("Submit")) then
        'first time entering page
        msg = ""
else
        if isempty(Request.Form("FirstName")) or Request.Form("FirstName") = "" then
                msg = "First Name is required."
        elseif isempty(Request.Form("LastName")) or Request.Form("LastName") = "" then
                msg = "Last Name is required."
        elseif isempty(Request.Form("EmailAddress")) or Request.Form("EmailAddress") = "" then
                msg = "Email is required."
        elseif isempty(Request.Form("Password")) or Request.Form("Password") = "" then
                msg = "Password is required."
        elseif isempty(Request.Form("RepeatPassword")) or Request.Form("RepeatPassword") = "" then
                msg = "Repeat Password is required."
        elseif Request.Form("Password") <> Request.Form("RepeatPassword") then
                msg = "Both passwords do not match!"
        else
                set RSStudent = conn.execute ("SELECT StudentID FROM students where StudentEmailAddress = '"_
                 & Request.Form("EmailAddress") & "'")
                if RSStudent.EOF then
                        'login available
                        set RSStudent = conn.execute ("INSERT INTO students (StudentFirstName, StudentLastName, StudentEmailAddress," _
                         & " StudentPassword, DateCreated) values ('" & Request.Form("FirstName") & "', '" & _
                         Request.Form("LastName") & "', '" & Request.Form("EmailAddress") & "', '" & Request.Form("Password") & "', '" & _
						 						 Now() & "')")
                        
                        set RSStudent = conn.execute ("SELECT StudentID FROM students where StudentEmailAddress = '"_
                         & Request.Form("EmailAddress") & "' and StudentPassword = '" _
                         & Request.Form("Password") & "'")
                        Session("StudentID") = RSStudent("StudentID")
                        


'-- ****************************************************************************      
'--this is new addition to create record upon initial login

				set RSTestsSelect = conn.execute ("SELECT * FROM followup_tests where StudentID=" & Session("StudentID") & "")
                
                 if RSTestsSelect.EOF then
                        set RSTestsInsert = conn.execute ("INSERT INTO followup_tests (StudentID, TestDate) values ('" & Session("StudentID") & "', '" & Date & "')")   
                else
					Session.Timeout=30
					Session("Question1") = RSTestsSelect("Question1")
					Session("Question2") = RSTestsSelect("Question2")
					Session("Question3") = RSTestsSelect("Question3")
					Session("Question4") = RSTestsSelect("Question4")
					Session("Question5") = RSTestsSelect("Question5")
					Session("Question6") = RSTestsSelect("Question6")
					Session("Question7") = RSTestsSelect("Question7")
					Session("Question8") = RSTestsSelect("Question8")
					Session("Question9") = RSTestsSelect("Question9")
					Session("Question10") = RSTestsSelect("Question10")
					Session("Wordiness1") = RSTestsSelect("Wordiness1")
					Session("Wordy2Rate") = RSTestsSelect("Wordy2Rate")
					Session("Wordy3Rate") = RSTestsSelect("Wordy3Rate")
					Session("Wordy4Rate") = RSTestsSelect("Wordy4Rate")
					Session("Long1Rate") = RSTestsSelect("Long1Rate")
					Session("Long2Rate") = RSTestsSelect("Long2Rate")

                end if 
'-- ****************************************************************************

Response.Redirect "./begin.asp"



                else
                        'login taken
                        msg = "The Email Address you've entered is in use."
                end if          
        end if
end if                  
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 TRANSITIONAL//EN">
<html>
<head>
<title>L'Oreal Writing Skills Assessment: Register</title>
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<META NAME="rating" CONTENT="General">
<META NAME="revisit-after" CONTENT="15 days">
<META HTTP-EQUIV = "expires" CONTENT="Never">
<META NAME="ROBOTS" CONTENT="ALL">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META name="description" content="Effective Business Writing Seminar: Register">
<META name="keywords" Content="writing at work effective business writing grammar puncutation">
<link rel="stylesheet" type="text/css" href="../unilever.css">
</head>

<body>

  <form action="./register.asp"  method="post">
<!-- This table sets the format for the page -->
  <table cellpadding="0" cellspacing="0" border="0" width="720" cols="3">

   <tr valign="top" align="left">
   
	<td width="170"></td>
   <!-- This column inserts a spacer -->	
	<td width="10">
	<img src="../assets/spacer.gif" width="10" border="0"></td>
	
    <td width="540"><img src="../assets/logo_loreal_big.gif" width="162" height="72" border="0">




		<h1>Registration</h1>
  
        <!-- This part prints the appropriate message -->
         <% if msg = "" Then %><h2>    
    Please fill in all the information below.</h2><h3> <%= msg %></h3>
    <% else %><h2>
    Information supplied is not correct.</h2><p class="red"> <%= msg %></p>
    <%end if %>
  
        <!-- This embedded table sets the form text boxes and the new student icon-->
        <table cellpadding="0" cellspacing="0" border="0" width="540" cols="2">
        <tr valign="top" align="left">
        <td width="100">
        <p>First Name<br>
        <input type="text" name="FirstName" size="30" maxlength="50" value="<% =Session("FirstName") %>" > </p> 
        <p>Last Name<br>
        <input type="text" name="LastName" size="30" maxlength="50" value="<% =Session("LastName") %>" > </p>   
        

        <p>Email Address<br>
        <input type="text" name="EmailAddress" size="30" maxlength="50" value="<% =Session("EmailAddress") %>" ></p>    
        <p>Password<br>
        <input type="password" name="Password" size="20" maxlength="20" value="<% =Session("Password") %>" ></p>        
        <p>Re-enter Password<br>
        <input type="password" name="RepeatPassword" size="20" maxlength="20" value="<% =Session("RepeatPassword") %>" ></p></td>
        <td width="100" align="center"><img src="../assets/pencil2.gif" width="50" height="50" border="0"><br>
        </td>
        </tr>
        </table>

        <!-- This part sets the submit and reset buttons -->
        <p>
        <input type="submit" name="Submit" value="Submit">      <input type="reset" value="Clear"> 
        </p>

        </td>
        </tr>
<!-- This closes out the row  -->

        </td>
        </tr>


<!-- This closes out the table -->
        </table>


<%
'-- This closes the recordsets and database connection
conn.Close
Set conn = Nothing
%>
</form>

</body>
</html>
