Option Explicit
Sub OutlookAutoGreeting()

    Dim origEmail As MailItem: Set origEmail = ActiveExplorer.Selection(1)
    Dim replyEmail As MailItem: Set replyEmail = origEmail.ReplyAll
    Dim SenderName1 as String
    Dim SenderName2 as String
            
    'Get current time
    Dim LHour: LHour = Hour(Now)
    
    'Pull SenderName from origEmail
    Dim SenderName: SenderName = Split(origEmail.SenderName)(0)
    
    'Ignore SenderName original case, make first char uppercase, all others lowercase
    SenderName1 = LCase(Right(SenderName, (Len(SenderName) - 1)))
    SenderName2 = UCase(Left(SenderName, 1))
    SenderName = SenderName2 & SenderName1
    
    'Generate time-dependent salutation
	Dim Morning As String = "Good morning "
	Dim Afternoon As String = "Good afternoon "
	Dim Evening As String = "Good evening "
	Dim TimeOfDay As String
    
    If (LHour <= 11) Then
        TimeOfDay = Morning
    ElseIf (LHour <= 16) Then
        TimeOfDay = Afternoon
    Else
        TimeOfDay = Evening
    End If
    
    'Append salutation with name
    Dim Greeting As String: Greeting = TimeOfDay & SenderName & ","
    
    'Assemble/display email content
    replyEmail.HTMLBody = Greeting & vbNewLine & replyEmail.HTMLBody & origEmail.Reply.HTMLBody
    replyEmail.Display

End Sub
