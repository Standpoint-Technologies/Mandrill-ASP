<!-- #include file="aspJSON.asp" -->
<%
  Const mandrillEmailAddressTo = "to"
  Const mandrillEmailAddressCC = "cc"
  Const mandrillEmailAddressBCC = "bcc"

  Const mandrillEmailStatusSent = "sent"
  Const mandrillEmailStatusQueued = "queued"
  Const mandrillEmailStatusScheduled = "scheduled"
  Const mandrillEmailStatusRejected = "rejected"
  Const mandrillEmailStatusInvalid = "invalid"

  Class MandrillApi
    Private m_ApiKey
    Public Property Get ApiKey()
      ApiKey = m_ApiKey
    End Property
    Public Property Let ApiKey(k)
      m_ApiKey = k
    End Property

    Public Function SendMessage(message)
      Dim postBody
      postBody = message.ToJSON(ApiKey)

      Dim jsonResponse
      jsonResponse = postJSON("messages/send.json", postBody)

      SendMessage = parseMandrillEmailResponse(jsonResponse)
    End Function

    Private Function postJSON(route, json)
        Dim o
        Set o = CreateObject("MSXML2.XMLHTTP")
        o.Open "POST", "https://mandrillapp.com/api/1.0/" + route, False
        o.SetRequestHeader "Content-Type", "application/json"
        o.SetRequestHeader "Content-Length", Len(json)
        o.Send json

        postJSON = o.ResponseText
    End Function

    Private Function parseMandrillEmailResponse(json)
      Dim i, j, r()
      Set j = New aspJSON
      j.loadJSON(json)

      ReDim r(j.data.Count)
      For Each i In j.data
        With j.data(i)
          Set r(i) = New MandrillEmailResponse
          r(i).Email = .Item("email")
          r(i).Status = .Item("status")
          r(i).RejectReason = .Item("reject_reason")
          r(i).ID = .Item("_id")
        End With
      Next

      parseMandrillEmailResponse = r
    End Function
  End Class

  Class MandrillEmailMessage
    Private m_PreserveRecipients
    Public Property Get PreserveRecipients()
      PreserveRecipients = m_PreserveRecipients
    End Property
    Public Property Let PreserveRecipients(p)
      m_PreserveRecipients = p
    End Property

    Private m_Subject
    Public Property Get Subject()
      Subject = m_Subject
    End Property
    Public Property Let Subject(s)
      m_Subject = s
    End Property

    Private m_Text
    Public Property Get Text()
      Text = m_Text
    End Property
    Public Property Let Text(t)
      m_Text = t
    End Property

    Private m_FromEmail
    Public Property Get FromEmail()
      FromEmail = m_FromEmail
    End Property
    Public Property Let FromEmail(f)
      m_FromEmail = f
    End Property

    Private m_Tags
    Public Property Get Tags()
      Tags = m_Tags
    End Property
    Public Property Let Tags(t)
      m_Tags = t
    End Property

    Private m_Recipients
    Public Property Get Recipients()
      Recipients = m_Recipients
    End Property
    Public Property Let Recipients(r)
      m_Recipients = r
    End Property

    Private m_HTML
    Public Property Get HTML()
      HTML = m_HTML
    End Property
    Public Property Let HTML(h)
      m_HTML = h
    End Property

    Private m_Attachments
    Public Property Get Attachments()
      Attachments = m_Attachments
    End Property

    Private m_Subaccount
    Public Property Get Subaccount()
      Subaccount = m_Subaccount
    End Property
    Public Property Let Subaccount(s)
      m_Subaccount = s
    End Property

    Private m_FromName
    Public Property Get FromName()
      FromName = m_FromName
    End Property
    Public Property Let FromName(f)
      m_FromName = f
    End Property

    Private m_Metadata
    Public Property Get Metadata()
      Metadata = m_Metadata
    End Property
    Public Property Let Metadata(m)
      Set m_Metadata = m
    End Property

    Private m_Headers
    Public Property Get Headers()
      Headers = m_Headers
    End Property
    Public Property Let Headers(h)
      Set m_Headers = h
    End Property

    Public Sub Class_Initialize
      Dim a()
      Redim a(0)
      m_Attachments = a
    End Sub

    Public Sub AddAttachment(attachment)
      ReDim Preserve m_Attachments(UBound(m_Attachments) + 1)
      Set m_Attachments(UBound(m_Attachments) - 1) = attachment
    End Sub

    Public Function ToJSON(apiKey)
      Dim json, i, m, h
      Set json = New aspJSON

      With json.data
        .Add "key", apiKey
        .Add "message", json.Collection()
        With json.data("message")
          .Add "text", m_Text
          .Add "html", m_HTML
          .Add "subject", m_Subject

          If Len(m_Subaccount) > 0 Then
            .Add "subaccount", m_Subaccount
          End If

          If Len(m_FromName) > 0 Then
            .Add "from_name", m_FromName
          End If

          If Len(m_FromEmail) > 0 Then
            .Add "from_email", m_FromEmail
          End If

          .Add "to", json.Collection()
          If IsArray(m_Recipients) Then
            With .Item("to")
              For i = 0 To UBound(m_Recipients) - 1
                If IsObject(m_Recipients(i)) Then
                  .Add i, json.Collection()
                  With .Item(i)
                    .Add "email", m_Recipients(i).Email
                    .Add "name", m_Recipients(i).Name
                    .Add "type", m_Recipients(i).AddressType
                  End With
                End If
              Next
            End With
          End If

          If IsArray(m_Tags) Then
            .Add "tags", json.Collection()
            With .Item("tags")
              For i = 0 To UBound(m_Tags) - 1
                .Add i, m_Tags(i)
              Next
            End With
          End If

          If IsArray(m_Attachments) Then
            .Add "attachments", json.Collection()
            With .Item("attachments")
              For i = 0 To UBound(m_Attachments) - 1
                If IsObject(m_Attachments(i)) Then
                  .Add i, json.Collection()
                  With .Item(i)
                    .Add "type", m_Attachments(i).AttachmentType
                    .Add "name", m_Attachments(i).Name
                    .Add "content", m_Attachments(i).Content
                  End With
                End If
              Next
            End With
          End If

          If StrComp(TypeName(m_Metadata), "Dictionary") = 0 Then
            .Add "metadata", json.Collection()
            With .Item("metadata")
              For Each m In m_Metadata
                .Add m, m_Metadata(m)
              Next
            End With
          End If

          If StrComp(TypeName(m_Headers), "Dictionary") = 0 Then
            .Add "headers", json.Collection()
            With .Item("headers")
              For Each h In m_Headers
                .Add h, m_Headers(h)
              Next
            End With
          End If
        End With
      End With

      ToJSON = json.JSONoutput()
    End Function
  End Class

  Class MandrillAttachment
    Private m_AttachmentType
    Public Property Get AttachmentType()
      AttachmentType = m_AttachmentType
    End Property
    Public Property Let AttachmentType(t)
      m_AttachmentType = t
    End Property

    Private m_Name
    Public Property Get Name()
      Name = m_Name
    End Property
    Public Property Let Name(n)
      m_Name = n
    End Property

    Private m_Content
    Public Property Get Content()
      Content = m_Content
    End Property
    Public Property Let Content(c)
      m_Content = c
    End Property
  End Class

  Class MandrillEmailAddress
    Private m_Email
    Public Property Get Email()
      Email = m_Email
    End Property
    Public Property Let Email(e)
      m_Email = e
    End Property

    Private m_Name
    Public Property Get Name()
      Name = m_Name
    End Property
    Public Property Let Name(n)
      m_Name = n
    End Property

    Private m_Type
    Public Property Get AddressType()
      AddressType = m_Type
    End Property
    Public Property Let AddressType(t)
      m_Type = t
    End Property
  End Class

  Class MandrillEmailResponse
    Private m_Email
    Public Property Get Email()
      Email = m_Email
    End Property
    Public Property Let Email(e)
      m_Email = e
    End Property

    Private m_Status
    Public Property Get Status()
      Status = m_Status
    End Property
    Public Property Let Status(s)
      m_Status = s
    End Property

    Private m_RejectReason
    Public Property Get RejectReason()
      RejectReason = m_RejectReason
    End Property
    Public Property Let RejectReason(r)
      m_RejectReason = r
    End Property

    Private m_ID
    Public Property Get ID()
      ID = m_ID
    End Property
    Public Property Let ID(i)
      m_ID = i
    End Property
  End Class
%>
