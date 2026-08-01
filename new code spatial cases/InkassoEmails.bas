Attribute VB_Name = "InkassoEmails"
Option Explicit

Sub ErstelleInkassoEmails()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim folderPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim kundenNr As String

    Dim merchID As String
    Dim emailAdresse As String
    Dim sprache As String
    Dim Betreff As String

    Dim OutlookApp As Object
    Dim mail As Object

    Dim heute As Date
    Dim i As Long

    heute = Date

    ' Folder where the Inkasso files are stored
    folderPath = "C:\Users\Desi\OneDrive - Zaala AG\Desislava @ Zaala AG - Linked Files\Inkasso Excel Table\UebergabenForPraxen\"

    ' Reference to the customer table
    Set ws = ThisWorkbook.Worksheets("Kunden")
    Set tbl = ws.ListObjects("tblKunden")

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' Open Outlook application
    Set OutlookApp = CreateObject("Outlook.Application")

    ' Check every file in the folder
    For Each file In folder.Files

        ' Check if file was modified today
        If DateValue(file.DateLastModified) = Date Then

            ' Extract the first 5 digits from the file name (customer number)
            kundenNr = Trim(Left(file.Name, 5))

            ' Search customer number in tblKunden MerchID column
            For i = 1 To tbl.ListRows.Count

                merchID = Trim(CStr(tbl.DataBodyRange(i, 6).Value))

                If merchID = kundenNr Then

                    ' Get customer email address and language
                    emailAdresse = Trim(tbl.DataBodyRange(i, 12).Value) ' Column L: eMailfuerDokumente
                    sprache = Trim(tbl.DataBodyRange(i, 24).Value)      ' Column X: Sprache

                    If emailAdresse <> "" Then

                        ' Build email subject based on customer language
                        If UCase(sprache) = "F" Then

                            Betreff = kundenNr & " Transfert a l'agence de recouvrement " & Format(heute, "dd.mm.yyyy")

                        ElseIf UCase(sprache) = "D" Then

                            Betreff = kundenNr & " Inkassoubergabe " & Format(heute, "dd.mm.yyyy")

                        Else

                            GoTo CustomerProcessed

                        End If


                        ' Check whether the same email has already been sent during the last 30 days
                        If MailAlreadySent(OutlookApp, emailAdresse, Betreff) Then

                            MsgBox "An email with the subject:" & vbCrLf & _
                                   Betreff & vbCrLf & vbCrLf & _
                                   "has already been sent to:" & vbCrLf & _
                                   emailAdresse, _
                                   vbExclamation, "Email already sent"

                            GoTo NextFile

                        End If


                        ' Create a new Outlook email
                        Set mail = OutlookApp.CreateItem(0)

                        With mail

                            ' Set sender mailbox
                            .SentOnBehalfOfName = "abrechnung@zaala.ch"

                            ' Set recipient
                            .To = emailAdresse

                            ' Set subject
                            .Subject = Betreff

                            ' Add Inkasso file as attachment
                            .Attachments.Add file.Path


                            If IsSpecialCustomer(kundenNr) Then

                                ' French email
                                If UCase(sprache) = "F" Then

                                    .Body = _
                                    "Bonjour," & vbCrLf & vbCrLf & _
                                    "Veuillez trouver ci-joint la liste des transferts de collecte d'aujourd'hui." & vbCrLf & vbCrLf & _
                                    "N'h" & ChrW(233) & "sitez pas " & ChrW(224) & " nous contacter si vous avez des questions." & vbCrLf & vbCrLf & _
                                    "Cordialement," & vbCrLf & vbCrLf & _
                                    "Centre de service clients" & vbCrLf & vbCrLf & _
                                    "Zaala AG" & vbCrLf & _
                                    "Bahnhofstrasse 51" & vbCrLf & _
                                    "8500 Frauenfeld" & vbCrLf & _
                                    "abrechnung@zaala.ch" & vbCrLf & _
                                    "www.zaala.ch"


                                ' German email
                                ElseIf UCase(sprache) = "D" Then

                                    .Body = _
                                    "Sehr geehrte Damen und Herren" & vbCrLf & vbCrLf & _
                                    "In der Beilage finden Sie die Auflistung der heutigen Inkasso" & ChrW(252) & "bergaben." & vbCrLf & vbCrLf & _
                                    "F" & ChrW(252) & "r Fragen stehen wir Ihnen gerne zur Verf" & ChrW(252) & "gung." & vbCrLf & vbCrLf & _
                                    "Freundliche Gr" & ChrW(252) & "sse" & vbCrLf & vbCrLf & _
                                    "Kundendienst Zaala AG" & vbCrLf & vbCrLf & _
                                    "Zaala AG" & vbCrLf & _
                                    "Bahnhofstrasse 51" & vbCrLf & _
                                    "8500 Frauenfeld" & vbCrLf & _
                                    "abrechnung@zaala.ch" & vbCrLf & _
                                    "www.zaala.ch"

                                End If

                            Else

                            ' French email
                            If UCase(sprache) = "F" Then

                                .Body = _
                                "Bonjour," & vbCrLf & vbCrLf & _
                                "Veuillez trouver ci-joint la liste des transferts de collecte d'aujourd'hui." & vbCrLf & _
                                "La r" & ChrW(233) & "trofacturation correspondante sera visible lors du prochain paiement." & vbCrLf & _
                                "N'h" & ChrW(233) & "sitez pas " & ChrW(224) & " nous contacter si vous avez des questions." & vbCrLf & vbCrLf & _
                                "Cordialement," & vbCrLf & _
                                "Centre de service clients" & vbCrLf & vbCrLf & _
                                "Zaala AG" & vbCrLf & _
                                "Bahnhofstrasse 51" & vbCrLf & _
                                "8500 Frauenfeld" & vbCrLf & _
                                "abrechnung@zaala.ch" & vbCrLf & _
                                "www.zaala.ch"


                            ' German email
                            ElseIf UCase(sprache) = "D" Then

                                .Body = _
                                "Sehr geehrte Damen und Herren" & vbCrLf & vbCrLf & _
                                "In der Beilage finden Sie die Auflistung der heutigen Inkasso" & ChrW(252) & "bergaben." & vbCrLf & _
                                "Die entsprechende R" & ChrW(252) & "ckbelastung wird bei der n" & ChrW(228) & "chsten Auszahlung ersichtlich sein." & vbCrLf & _
                                "F" & ChrW(252) & "r Fragen stehen wir Ihnen gerne zur Verf" & ChrW(252) & "gung." & vbCrLf & vbCrLf & _
                                "Freundliche Gr" & ChrW(252) & "sse" & vbCrLf & _
                                "Kundendienst Zaala AG" & vbCrLf & vbCrLf & _
                                "Zaala AG" & vbCrLf & _
                                "Bahnhofstrasse 51" & vbCrLf & _
                                "8500 Frauenfeld" & vbCrLf & _
                                "abrechnung@zaala.ch" & vbCrLf & _
                                "www.zaala.ch"

                            End If

                            End If


                            ' Save email as draft
                            .Save

                        End With

                    End If


CustomerProcessed:

                    ' Stop searching after customer was found
                    Exit For

                End If

            Next i

        End If


NextFile:

    Next file


    MsgBox "Inkasso emails have been created as drafts.", vbInformation


End Sub



Private Function IsSpecialCustomer(ByVal KundenNr As String) As Boolean

    Select Case KundenNr
        Case "85580", "85585", "85590", _
             "85661", "85662", "85666", "85667", _
             "85670", "85671", _
             "85935", "85945"

            IsSpecialCustomer = True
            Exit Function

    End Select

    If Left$(KundenNr, 2) = "95" Then
        IsSpecialCustomer = True
        Exit Function
    End If

    IsSpecialCustomer = False

End Function



Private Function MailAlreadySent(ByVal OutlookApp As Object, _
                                 ByVal Recipient As String, _
                                 ByVal SubjectText As String) As Boolean

    Dim Ns As Object
    Dim SharedMailbox As Object
    Dim SentFolder As Object
    Dim Items As Object
    Dim FilteredItems As Object
    Dim itm As Object
    Dim FilterDate As Date
    Dim Filter As String


    ' Get Outlook namespace
    Set Ns = OutlookApp.GetNamespace("MAPI")


    ' Get shared mailbox
    On Error Resume Next
    Set SharedMailbox = Ns.Folders("Zaala Service Center Kunden")
    On Error GoTo 0


    If SharedMailbox Is Nothing Then

        MsgBox "Shared mailbox 'Zaala Service Center Kunden' was not found in Outlook.", vbCritical

        MailAlreadySent = False
        Exit Function

    End If


    ' Get Sent Items folder of shared mailbox
    On Error Resume Next
    Set SentFolder = SharedMailbox.Folders("Gesendete Elemente")
    On Error GoTo 0


    If SentFolder Is Nothing Then

        MsgBox "Folder 'Gesendete Elemente' was not found in shared mailbox.", vbCritical

        MailAlreadySent = False
        Exit Function

    End If


    ' Search only emails sent during the last 30 days
    FilterDate = Date - 30


    Set Items = SentFolder.Items


    ' Sort emails by sent date descending
    Items.Sort "[SentOn]", True


    ' Build Outlook filter
    Filter = "[SentOn] >= '" & Format(FilterDate, "ddddd h:nn AMPM") & "'"


    Set FilteredItems = Items.Restrict(Filter)


    ' Check matching emails
    For Each itm In FilteredItems


        If itm.Class = 43 Then


            If LCase(Trim(itm.Subject)) = LCase(Trim(SubjectText)) Then


                If InStr(1, LCase(itm.To), LCase(Recipient), vbTextCompare) > 0 Then


                    MailAlreadySent = True
                    Exit Function


                End If


            End If


        End If


    Next itm


    MailAlreadySent = False


End Function

