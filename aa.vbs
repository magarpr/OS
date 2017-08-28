WScript.Echo EMail( "Rajeev <from@from.com>", _
                    "Prashant Magar <prashant.magar@in.fujitsu.com>", _
                    "Hi", _
                    "<html><body><div><br/>Dear how are you </div></body></html>", _
                    "<html><body><div><br/><a href='https://www.facebook.com/'>open  form </a> </div></body></html>", _
                    "", _
                    "59.160.188.194", _
                    25 ,"domain\user","pass")

Function EMail( myFrom, myTo, mySubject, myTextBody, myHTMLBody, myAttachment, mySMTPServer, mySMTPPort,user,pass )
 
    Dim i, objEmail

    
    On Error Resume Next

     
    Set objEmail = CreateObject( "CDO.Message" )

     
    With objEmail
        .From     = myFrom
        .To       = myTo
        ' Other options you might want to add:
        ' .Cc     = ...
        ' .Bcc    = ...
        .Subject  = mySubject
        .TextBody = myTextBody
        .HTMLBody = myHTMLBody
        If IsArray( myAttachment ) Then
            For i = 0 To UBound( myAttachment )
                .AddAttachment Replace( myAttachment( i ), "\", "\\" ),"",""
            Next
        ElseIf myAttachment <> "" Then
            .AddAttachment Replace( myAttachment, "\", "\\" ),"",""
        End If
        If mySMTPPort = "" Then
            mySMTPPort = 25
        End If
        With .Configuration.Fields
            .Item( "http://schemas.microsoft.com/cdo/configuration/sendusing"      ) = 2
            .Item( "http://schemas.microsoft.com/cdo/configuration/smtpserver"     ) = mySMTPServer
            .Item( "http://schemas.microsoft.com/cdo/configuration/smtpserverport" ) = mySMTPPort
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")    = user
            .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")    = pass
            .Update
        End With
        ' Send the message
        .Send
    End With
    ' Return status message
    If Err Then
        EMail = "ERROR " & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        EMail = "Message sent ok"
    End If

    ' Release the e-mail message object
    Set objEmail = Nothing
    ' Restore default error handling
    On Error Goto 0
End Function