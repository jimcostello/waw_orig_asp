<html>
<head><title>ASP Mailer Test from Aries</title><head>
<body>
<H3>ASP Mailer Test from Aries</H3>
<%

  rem ------------------------------------
  rem - Instantiate the SMTP mailer object
  rem ------------------------------------
  Set Mailer = Server.CreateObject("SMTPsvg.Mailer")


  rem ------------------------------------
  rem - Get the internal version number
  rem ------------------------------------
  rem strVer = Mailer.Version

  rem ------------------------------------
  rem - Set the FromName and FromAddress
  rem - WARNING: These are REQUIRED FIELDS
  rem ------------------------------------
  Mailer.FromName = "support"
  Mailer.FromAddress = "support@1stchoicesite.com"

  rem ------------------------------------
  rem - Set the Remote Host (SMTP) Host
  rem -   that we'll send this mail thru
  rem ------------------------------------
  strMailHost = "Capricorn.1stchoice-nt.net"
  if strMailHost = "gemini.1stchoice-nt.net" then
    Response.Write "<h2 align=center>Error: You need to change the mailhost to your SMTP server address</h2>"
  end if
  Mailer.RemoteHost = strMailHost

  rem ------------------------------------
  rem - Set the recipient's address
  rem -   that this message will go to
  rem ------------------------------------
  Mailer.AddRecipient "Support", "support@1stchoicesite.com"

  rem ------------------------------------
  rem - Set the Carbon Copy addresses
  rem ------------------------------------
  rem Mailer.AddCC "Steve", "xyz@abcxyz.net"

  rem ------------------------------------
  rem - Set the Blind Carbon Copy addresses
  rem ------------------------------------
  rem Mailer.AddBCC "George Higgins", "ghi@abcxyz.net"

  rem ------------------------------------
  rem - Set the ReturnReceipt flag
  rem -   If this is set to true AND the
  rem -   recipients SMTP server supports
  rem -   this feature (and it is enabled)
  rem -   the recipients SMTP server will
  rem -   send a notice back to the FromAddress
  rem -   confirming that this email has been
  rem -   delivered.
  rem ------------------------------------
  Mailer.ReturnReceipt = false

  rem ------------------------------------
  rem - Set the ConfirmReading flag
  rem -   If this is set to true AND the
  rem -   recipients email program supports
  rem -   this feature (and it is enabled)
  rem -   the recipients email program will
  rem -   send a notice back to the FromAddress
  rem -   confirming that this email has been
  rem -   read.
  rem ------------------------------------
  Mailer.ConfirmRead = false

  rem ------------------------------------
  rem - Set the subject line
  rem ------------------------------------
  Mailer.Subject = "ASP Test Successful!"

  rem ------------------------------------
  rem - Set the message body text
  rem -
  rem - To optionally clear the text once
  rem - you have set it use the ClearBodyText
  rem - method (e.g. Mailer.ClearBodyText
  rem ------------------------------------
  strBody = "Dear Stephen" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
  strBody = strBody & "This is really a cool product. I'd be happy "
  strBody = strBody & "to pay a small fortune for such a nifty product." & Chr(13) & Chr(10) & Chr(13) & Chr(10)
  strBody = strBody & "Sincerely," & Chr(13) & Chr(10) & Chr(13) & Chr(10)
  strBody = strBody & "Happy ASP Programmer from Aries(" & Now & ")"
  Mailer.BodyText = strBody


  rem ------------------------------------
  rem - Clear the attachments and add a
  rem -  couple of files. Make sure that
  rem -  the IUSR_XYZ IIS user has security
  rem -  rights that allow the component
  rem -  to read the necessary files!
  rem -  (optional)
  rem - Attachments are not supported in
  rem -  the eval version
  rem ------------------------------------
  rem Mailer.ClearAttachments
  rem Mailer.AddAttachment "c:\autoexec.bat"
  rem Mailer.AddAttachment "c:\config.sys"

  rem ------------------------------------
  rem - Set the encoding type (default is MIME)
  rem -   1 = UUEncoded
  rem -   2 = MIME
  rem - (optional)
  rem ------------------------------------
  rem Mailer.Encoding = 2

  rem ------------------------------------
  rem - If you need to debug the session
  rem -  give a log file name here. Make
  rem -  sure the IUSR_XYZ IIS user has
  rem -  security that allows the component
  rem -  to write to this file
  rem -  (optional)
  rem -
  rem - *** WARNING *** *** WARNING ***
  rem - do not use this setting in situations
  rem - where multiple users can acess this
  rem - component at the same time. This is
  rem - for single user debugging ONLY!
  rem - *** WARNING *** *** WARNING ***
  rem ------------------------------------
  rem Mailer.SMTPLog = "c:\smtp.log"

  rem ------------------------------------
  rem - Set priority and if message is
  rem -  urgent. Prioritys are 1, 3, 5 and
  rem -  are reflected in the X-Priority
  rem -  header
  rem - 1 = High, 3 = Normal, 5 = Low
  rem -  (optional)
  rem ------------------------------------
  rem Mailer.Priority = 1
  rem Mailer.Urgent = false

  rem ------------------------------------
  rem - Set the character set
  rem -  By default the char set is US Ascii
  rem -
  rem -  Valid values:
  rem -  1 = US Ascii
  rem -  2 = ISO-8859-1
  rem ------------------------------------
  rem Mailer.CharSet = 1

  rem ------------------------------------
  rem - Set the TimeOut
  rem   Default timeout is 60
  rem ------------------------------------
  rem Mailer.Timeout = 30

  rem ------------------------------------
  rem - Send the message
  rem ------------------------------------
  rem SendMail returns a true or false
  rem  success result you can optionally
  rem  test for
  if not Mailer.SendMail then
    if Mailer.Response <> "Mail Sent" then
      strError = Mailer.Response
    else
      strError = "Unknown"
    end if
    Response.Write "Mail failure occured. Reason: " & strError
end if


%>

</body>
</html>