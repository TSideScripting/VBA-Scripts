Option Explicit

Public Sub EmailAlert()
    Dim Trkr As Worksheet
    Dim McrS As Worksheet
    Dim mApp As Object
    Dim mMail As Object
    Dim SendToMail As String
    Dim CCMail as String
    Dim MailSubject As String
    Dim mMailBody As String
    Dim Eqpt(0 To 10) As Long
    Dim EqptName(0 To 10) As String
    Dim Status(0 To 10) As String
    Dim Stat0(0 To 10) As String
    Dim Email(0 To 10) As String
    Dim Email0(0 To 10) As String
    Dim i As Long

    Set Trkr = ThisWorkbook.Worksheets("Tracker")
    Set McrS = ThisWorkbook.Worksheets("MacroStuff")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = 0 To 10

        Eqpt(i) = Trkr.Range("C1").Offset(0, i).Value
        EqptName(i) = Trkr.Range("C2").Offset(0, i).Value

        If Eqpt(i) <= 20 And Eqpt(i) <> 0 Then
            Status(i) = True
        Else
            Status(i) = False
            McrS.Range("B2").Offset(i).Value = False
        End If
        If Eqpt(i) = 0 Then
            Stat0(i) = True
        Else
            Stat0(i) = False
            McrS.Range("C2").Offset(i).Value = False
        End If

        Email(i) = McrS.Range("B2").Offset(i).Value
        Email0(i) = McrS.Range("C2").Offset(i).Value

        If Status(i) = True And Email(i) = False Then
            If i <> 4 Then
                SendToMail = "ithelpdeskteam@uk.aswatson.com; tsinfrastructure@uk.aswatson.com"
                CCMail = ""
            Elseif i = 4 Then
                SendToMail = "ithelpdeskteam@uk.aswatson.com; Andre.Veredas@uk.aswatson.com"
                CCMail = "Mark.Garrett@uk.aswatson.com"
            End If
            MailSubject = "LOW STOCK ALERT - " & EqptName(i)
            mMailBody = EqptName(i) & "s are low on stock! Only " & Eqpt(i) & " units left!"
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
            McrS.Range("B2").Offset(i).Value = True
            Email(i) = True
        ElseIf Stat0(i) = True And Email0(i) = False Then
            If i <> 4 Then
                SendToMail = "ithelpdeskteam@uk.aswatson.com; tsinfrastructure@uk.aswatson.com"
                CCMail = ""
            Elseif i = 4 Then
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
            McrS.Range("C2").Offset(i).Value = True
            Email0(i) = True
        End If
    Next i
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


