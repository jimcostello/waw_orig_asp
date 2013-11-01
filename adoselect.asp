<HTML>
<% 
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% File: 	ADOselect.asp 
'% Author: 	Aaron L. Barth (MS)
'% Purpose: 	For testing ADO connectivity to any ODBC Datasource
'% Disclaimer: 	This code is to be used for sample purposes only
'%              Microsoft does not gaurantee its functionality
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

if Request("REQUESTTYPE") <> "POST" then
'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	% If the request does not contain REQUESTTYPE = "POST
'	% then display Form Page
'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

	dsn = Session("dsn")
	dbuser = Session("dbuser")
	dbpass = Session("dbpass")
	dbtable = Session("dbtable")
	dbfield = Session("dbfield")
	dbwhere = Session("dbwhere")
%>
	<FORM ACTION=adoselect.asp method=POST>
	<TABLE>
	<TR><TD><B>You are authenticated as: </TD>
		<TD><FONT COLOR=GREEN><% = Request.ServerVariables("LOGON_USER")%></TD></TR>
	<TR><TD><B>Your IP Address is: </TD>
		<TD><FONT COLOR=GREEN><% = Request.ServerVariables("REMOTE_ADDR")%></TD></TR>
	<TR><TD><B>System DSN:</TD> 
		<TD><INPUT TYPE=TEXT NAME=datasource VALUE="<% = dsn %>"></TD></TR>
	<TR><TD><B>Username:</TD> 
		<TD><INPUT TYPE=TEXT NAME=username VALUE="<% = dbuser %>"></TD></TR>
	<TR><TD><B>Password:</TD> 
		<TD><INPUT TYPE=Password NAME=password VALUE="<% = dbpass %>"></TD></TR>
	<TR><TD><B>Table:</TD> 
		<TD><INPUT TYPE=TEXT NAME=table VALUE="<% = dbtable %>"></TD></TR>
	<TR><TD><B><FONT COLOR=RED>WHERE</TD>
		<TD></TD></TR>
	<TR><TD><B>Field to Query:</TD> 
		<TD><INPUT TYPE=TEXT NAME=field VALUE="<% = dbfield %>"></TD></TR>
	<TR><TD><B>Value to Query:</TD> 
		<TD><INPUT TYPE=TEXT NAME=where VALUE="<% = dbwhere %>"></TD></TR>
	</TABLE>
	<INPUT TYPE=HIDDEN NAME=REQUESTTYPE VALUE="POST">
	<INPUT TYPE=Submit VALUE="Query Database">
	<HR>
	</FORM>
<%

else
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'% Perform Query to Database
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	% Request the datsource  from the Previous Form
'	% Set the Session variable so we can retrieve the 
'	% value for the next query
'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

	dsn  = Request("datasource")
	Session("dsn") = dsn

'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	% Request the username  from the Previous Form
'	% Set the Session variable so we can retrieve the 
'	% value for the next query
'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

	dbuser  = Request("username")
	Session("dbuser") = dbuser

'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	% Request the password from the Previous Form
'	% Set the Session variable so we can retrieve the 
'	% value for the next query
'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

	dbpass = Request("password")
	Session("dbpass") = dbpass

'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	% Request the table from the Previous Form
'	% Set the Session variable so we can retrieve the 
'	% value for the next query
'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

	dbtable = Request("table")
	Session("dbtable") = dbtable

'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	% Request the table from the Previous Form
'	% Set the Session variable so we can retrieve the 
'	% value for the next query
'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

	dbfield = Request("field")
	Session("dbfield") = dbfield

'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	% Request the table from the Previous Form
'	% Set the Session variable so we can retrieve the 
'	% value for the next query
'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

	dbwhere = Request("where")
	Session("dbwhere") = dbwhere

'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'	% Check to see if any of the requested values are blank, IF they 
'	% are, then inform the user which variables are blank ELSE
'	% Continue with the query
'	%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	if dsn = "" OR dbuser = "" OR dbtable = "" then

		Response.write "Error in SQL Statement:<BR>"
		if dsn = "" then
			Response.write "<FONT COLOR=RED>Missing System DSN</FONT><P>"
		end if
		if dbuser = "" then
			Response.write "<FONT COLOR=RED>Missing Username</FONT><P>"
		end if
		if dbtable = "" then
			Response.write "<FONT COLOR=RED>Missing Tablename</FONT><P>"
		end if
			Response.write "<FORM ACTION=adoselect.asp><INPUT TYPE=SUBMIT VALUE=ReQuery></FORM>"
	else
'		%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'		% Create the Conn Object and open it
'		% with the supplied parameters
'		% System DSN, UserID, Password
'		%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

		Set Conn = Server.CreateObject("ADODB.Connection")
		Set rs = Server.CreateObject("ADODB.RecordSet")
		Conn.Open dsn, dbuser, dbpass

'		%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'		% Build the SQL Statement and assign it
'		% to the variable sql.  Concatinating the dbtable and the SELECT
'		% statement
'		%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
		if dbfield = "" OR dbwhere ="" then
		sql="SELECT * FROM " & dbtable
		else

'		%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'		% IF dbfield and dbwhere are specified, then
'		% change the SQL statement to use the WHERE clause
'		%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'
		sql="SELECT * FROM " & dbtable
		sql = sql & " WHERE " & dbfield
		sql = sql & " LIKE '%" & dbwhere & "%'"
		end if

'		%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'		% For Debugging, Echo the SQL Statement
'		%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
		Response.Write "<B><FONT SIZE=2 COLOR=BLUE>SQL STATEMENT: </B>" & sql & "<HR>" 

'		%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'		% Open the RecordSet (RS) and pass it 
'		% the connection (conn) and the SQL Statement (sql)
'		%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
		RS.Open sql, Conn
		%>
		
		<P>
		<TABLE BORDER=1>
			<TR>
			<% 
'			%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'			% Loop through Fields Names and print out the Field Names
'			%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
		
			For i = 0 to RS.Fields.Count - 1 
			%>
			<TD><B><% = RS(i).Name %></B></TD>
			<% Next %>
			</TR>
			<% 
'			%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'			% Loop through rows, displaying each field
'			%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
			Do While Not RS.EOF 
			%>
			<TR>
			<% For i = 0 to RS.Fields.Count - 1 %>
			<TD VALIGN=TOP><% = RS(i) %></TD>
			<% Next %>
			</TR>
			<%
			RS.MoveNext
			Loop
'			%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'			% Make sure to close the Result Set and the Connection object
'			%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
			RS.Close
			Conn.Close
			%>
		</TABLE>

	<% 
	end if 
end if 
%>