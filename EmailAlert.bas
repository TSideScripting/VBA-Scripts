Option Explicit

Public Sub EmailAlert()
    Dim Trkr As Worksheet
    Dim McrS As Worksheet
    Dim mApp As Object
    Dim mMail As Object
    Dim SendToMail As String
    Dim CCMail As String
    Dim MailSubject As String
    Dim mMailBody As String
    Dim Eqpt(0 To 11) As Long
    Dim EqptName(0 To 11) As String
    Dim StatTwenty(0 To 11) As String
    Dim StatTen(0 To 11) As String
    Dim StatZero(0 To 11) As String
    Dim EmailTwenty(0 To 11) As String
    Dim EmailTen(0 To 11) As String
    Dim EmailZero(0 To 11) As String
    Dim i As Long

    Set Trkr = ThisWorkbook.Worksheets("Tracker")
    Set McrS = ThisWorkbook.Worksheets("MacroStuff")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = 0 To 11

        Eqpt(i) = Trkr.Range("C1").Offset(0, i).Value
        EqptName(i) = Trkr.Range("C2").Offset(0, i).Value

        If Eqpt(i) <= 20 And Eqpt(i) > 10 Then
            StatTwenty(i) = True
        Else
            StatTwenty(i) = False
            McrS.Range("B2").Offset(i).Value = False
        End If
        If Eqpt(i) <= 10 And Eqpt(i) <> 0 Then
            StatTen(i) = True
        Else
            StatTen(i) = False
            McrS.Range("C2").Offset(i).Value = False
        End If
        If Eqpt(i) = 0 Then
            StatZero(i) = True
        Else
            StatZero(i) = False
            McrS.Range("D2").Offset(i).Value = False
        End If

        EmailTwenty(i) = McrS.Range("B2").Offset(i).Value
        EmailTen(i) = McrS.Range("C2").Offset(i).Value
        EmailZero(i) = McrS.Range("D2").Offset(i).Value

        If StatTwenty(i) = True And EmailTwenty(i) = False Then
            If EqptName(i) <> "Fingerprint Scanner" Then
                SendToMail = "ithelpdeskteam@uk.aswatson.com; tsinfrastructure@uk.aswatson.com"
                CCMail = ""
            ElseIf EqptName(i) = "Fingerprint Scanner" Then
                SendToMail = "ithelpdeskteam@uk.aswatson.com; Andre.Veredas@uk.aswatson.com"
                CCMail = "Mark.Garrett@uk.aswatson.com"
            End If
            MailSubject = "LOW STOCK ALERT - " & EqptName(i)
            mMailBody = EqptName(i) & "s are running low on stock. There are currently " & Eqpt(i) & " units left."
            Set mApp = CreateObject("Outlook.Application")
            Set mMail = mApp.CreateItem(0)
            With mMail
                .To = SendToMail
                .CC = CCMail
                .Subject = MailSubject
                .Body = mMailBody
                .Importance = 1
                .Send ' You can use .Display
            End With
            McrS.Range("B2").Offset(i).Value = True
            EmailTwenty(i) = True
        ElseIf StatTen(i) = True And EmailTen(i) = False Then
            If EqptName(i) <> "Fingerprint Scanner" Then
                SendToMail = "ithelpdeskteam@uk.aswatson.com; tsinfrastructure@uk.aswatson.com"
                CCMail = ""
            ElseIf EqptName(i) = "Fingerprint Scanner" Then
                SendToMail = "ithelpdeskteam@uk.aswatson.com; Andre.Veredas@uk.aswatson.com"
                CCMail = "Mark.Garrett@uk.aswatson.com"
            End If
            MailSubject = "LOW STOCK ALERT - " & EqptName(i)
            mMailBody = EqptName(i) & "s are now low on stock! Only " & Eqpt(i) & " units left!"
            Set mApp = CreateObject("Outlook.Application")
            Set mMail = mApp.CreateItem(0)
            With mMail
                .To = SendToMail
                .CC = CCMail
                .Subject = MailSubject
                .Body = mMailBody
                .Importance = 2
                .Send ' You can use .Display
            End With
            McrS.Range("C2").Offset(i).Value = True
            EmailTen(i) = True
        ElseIf StatZero(i) = True And EmailZero(i) = False Then
            If EqptName(i) <> "Fingerprint Scanner" Then
                SendToMail = "ithelpdeskteam@uk.aswatson.com; tsinfrastructure@uk.aswatson.com"
                CCMail = ""
            ElseIf EqptName(i) = "Fingerprint Scanner" Then
                SendToMail = "ithelpdeskteam@uk.aswatson.com; Andre.Veredas@uk.aswatson.com"
                CCMail = "Mark.Garrett@uk.aswatson.com"
            End If
            MailSubject = "NO STOCK ALERT - " & EqptName(i)
            mMailBody = "THERE ARE CURRENTLY NO MORE " & UCase(EqptName(i)) & "S IN STOCK!"
            Set mApp = CreateObject("Outlook.Application")
            Set mMail = mApp.CreateItem(0)
            With mMail
                .To = SendToMail
                .CC = CCMail
                .Subject = MailSubject
                .Body = mMailBody
                .Importance = 2
                .Send ' You can use .Display
            End With
            McrS.Range("D2").Offset(i).Value = True
            EmailZero(i) = True
        End If
    Next i
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


Day-Walker18