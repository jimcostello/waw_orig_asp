<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

<title>ASP DSN Connection To A MySQL Database</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

<% 

Dim conn 

Set conn = Server.CreateObject("ADODB.Connection") 

conn.open "DSN=db_writingatworkonline_com;UID=writingatw160520;PWD=schultz" 

 

sqlquery = "SELECT * FROM student" 

Set rs = conn.Execute(sqlquery) 

'Response.Write "FieldName: <input type=text name="&TableName&" value="&FieldName&">" 

do while not rs.eof 

               response.write( "ASP DSN Connection To A MySQL Database: " & rs(0) ) 

               rs.movenext 

loop 

%>

76/0%
