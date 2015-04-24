<!-- #include file="Mandrill.asp" -->
<%
  Dim man, message, tags(1), recipients(3), metadata, headers, x
  Set man = New MandrillApi
  man.ApiKey = "<your-api-key>"

  Set message = New MandrillEmailMessage
  message.Text = "Test message body"
  message.HTML = "<b>Test HTML body</b>"
  message.Subject = "Test Subject"
  message.FromName = "Standpoint"
  message.FromEmail = "support@standpointtech.com"

  tags(0) = "test-tag"
  message.Tags = tags

  Set recipients(0) = New MandrillEmailAddress
  recipients(0).Email = "tbenfield@standpointtech.com"
  recipients(0).Name = "Tyler Benfield"
  recipients(0).AddressType = mandrillEmailAddressTo
  'Set recipients(1) = New MandrillEmailAddress
  'recipients(1).Email = "relting@standpointtech.com"
  'recipients(1).Name = "Robbie Elting"
  'recipients(1).AddressType = mandrillEmailAddressCC
  'Set recipients(2) = New MandrillEmailAddress
  'recipients(2).Email = "rtbenfield@gmail.com"
  'recipients(2).Name = "Tyler"
  'recipients(2).AddressType = mandrillEmailAddressBCC
  message.Recipients = recipients

  Set metadata = Server.CreateObject("Scripting.Dictionary")
  metadata("test-metadata") = "test"
  message.Metadata = metadata

  Set headers = Server.CreateObject("Scripting.Dictionary")
  headers("test") = "test-header"
  message.Headers = headers

  x = man.SendMessage(message)

  For i = 0 To UBound(x) - 1
    With x(i)
      If StrComp(.Status, mandrillEmailStatusSent) = 0 Or StrComp(.Status, mandrillEmailStatusQueued) = 0 Then
        Response.Write "Successfully sent email to " & .Email & " with response id " & .ID & " status " & .Status
      Else
        Response.Write "Error sending email to " & .Email & " status message " & .Status & " reject reason " & .RejectReason
      End If
    End With
  Next
%>
