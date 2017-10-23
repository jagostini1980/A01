Attribute VB_Name = "SendMail"
Option Explicit

Public poSendMail As New vbSendMail.clsSendMail

Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean


Public Sub EnviarEmail(Origen As String, Destino As String, Asunto As String, MSG As String, Optional Adjunto As String = "")
 'On Error GoTo ErrorEMail
 On Error Resume Next
    ' *****************************************************************************
    ' This is where all of the Components Properties are set / Methods called
    ' *****************************************************************************

    'cmdSend.Enabled = False
    'lstStatus.Clear
    'Screen.MousePointer = vbHourglass

    With poSendMail

        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************

        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = "mail.trcnet.com.ar" '"200.43.113.2"                 ' Required the fist time, optional thereafter
        .SMTPPort = 25
        
        .From = Origen                             ' Required the fist time, optional thereafter
        .FromDisplayName = Origen                  ' Optional, saved after first use
        .Recipient = Destino                       ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = Destino            ' Optional, separate multiple entries with delimiter character
        '.CcRecipient = txtCc                      ' Optional, separate multiple entries with delimiter character
        '.CcDisplayName = txtCcName                ' Optional, separate multiple entries with delimiter character
        '.BccRecipient = txtBcc                    ' Optional, separate multiple entries with delimiter character
        '.ReplyToAddress = txtFrom.Text            ' Optional, used when different than 'From' address
        .Subject = Asunto                  ' Optional
        .Message = MSG                      ' Optional
        .Attachment = Trim(Adjunto)          ' Optional, separate multiple entries with delimiter character

        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .UserName = "epq18"                         ' Optional, default = Null String
        .Password = "cumbia"                        ' Optional, default = Null String, value is NOT saved
        '.POP3Host = txtPopServer
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
        ' .SMTPPort = 25                            ' Optional, default = 25

        ' **************************************************************************
        ' OK, all of the properties are set, send the email...
        ' **************************************************************************
        ' .Connect                                  ' Optional, use when sending bulk mail
        '#### envia local ####
        .Send                                       ' Required
        ' .Disconnect                               ' Optional, use when sending bulk mail
        'txtServer.Text = .SMTPHost                 ' Optional, re-populate the Host in case
        '##### envia Remoto #####
        .SMTPPort = 2525                            ' MX look up was used to find a host    End With
        .Send
    End With
    'Screen.MousePointer = vbDefault
    'cmdSend.Enabled = True
'ErrorEMail:
'    Call ManipularError(Err.Number, Err.Description)
End Sub

