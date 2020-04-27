Sub Email()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody0 As String
	
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    strbody0 = "<body>This is the email text in HTML" & _
	"<br> The & _ is necessary for next HTML line </body>"

    On Error Resume Next

    With OutMail
        .Display
        .To = ""
        .CC = ""
        .Subject = ""
        .HTMLBody = strbody0
        .VotingOptions = "Agree; Disagree"
        .ReplyRecipients.Add ""
    End With

    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub



