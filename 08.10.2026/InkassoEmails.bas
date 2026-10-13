Attribute VB_Name = "InkassoEmails"
Option Explicit

Sub ErstelleInkassoEmails()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim folderPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim KundenNr As String

    Dim merchID As String
    Dim emailAdresse As String
    Dim sprache As String
    Dim Betreff As String

    Dim OutlookApp As Object
    Dim mail As Object

    Dim heute As Date
    Dim i As Long

    Dim createdCount As Long
    Dim skippedCount As Long
    Dim notFoundCount As Long
    Dim errorCount As Long

    On Error GoTo ErrorHandler

    heute = Date

    ' Folder where the Inkasso files are stored
    folderPath = "...\UebergabenForPraxen\"

    ' Reference to the customer table
    Set ws = ThisWorkbook.Worksheets("Kunden")
    Set tbl = ws.ListObjects("tblKunden")

    ' Check if the folder exists
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        MsgBox "The Inkasso folder was not found:" & vbCrLf & vbCrLf & _
               folderPath, vbCritical, "Folder not found"
        Exit Sub
    End If

    Set folder = fso.GetFolder(folderPath)

    ' Open Outlook application
    Set OutlookApp = CreateObject("Outlook.Application")

    ' Check every file in the folder
    For Each file In folder.Files

        On Error GoTo FileError

        ' Check if file was modified today
        If DateValue(file.DateLastModified) = heute Then

            ' Extract the first 5 digits from the file name
            KundenNr = Trim(Left(file.Name, 5))

            ' Search customer number in tblKunden
            For i = 1 To tbl.ListRows.Count

                ' Column E = Merch ID
                merchID = Trim(CStr(tbl.DataBodyRange(i, 5).Value))

                If merchID = KundenNr Then

                    ' Column G = eMail fur Dokumente
                    emailAdresse = Trim(CStr(tbl.DataBodyRange(i, 7).Value))

                    ' Column W = Sprache
                    sprache = Trim(CStr(tbl.DataBodyRange(i, 23).Value))

                    If emailAdresse <> "" Then

                        ' Build email subject based on customer language
                        If UCase(sprache) = "F" Then

                            Betreff = KundenNr & _
                                      " Transfert a l'agence de recouvrement " & _
                                      Format(heute, "dd.mm.yyyy")

                        ElseIf UCase(sprache) = "D" Then

                            Betreff = KundenNr & _
                                      " Inkassoubergabe " & _
                                      Format(heute, "dd.mm.yyyy")

                        Else

                            skippedCount = skippedCount + 1
                            GoTo CustomerProcessed

                        End If

                        ' Check whether the same email has already been sent
                        ' during the last 30 days
                        If MailAlreadySent(OutlookApp, emailAdresse, Betreff) Then

                            skippedCount = skippedCount + 1
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

                            If IsSpecialCustomer(KundenNr) Then

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
                                    "xxx@zaala.ch" & vbCrLf & _
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
                                    "xxx@zaala.ch" & vbCrLf & _
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
                                    "xxx@zaala.ch" & vbCrLf & _
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
                                    "xxx@zaala.ch" & vbCrLf & _
                                    "www.zaala.ch"

                                End If

                            End If

                            ' Save email as draft
                            .Save

                        End With

                        createdCount = createdCount + 1

                    Else

                        ' No email address in column G
                        skippedCount = skippedCount + 1

                    End If

CustomerProcessed:

                    ' Stop searching after customer was found
                    Exit For

                End If

            Next i

            ' Customer number was not found
            If i > tbl.ListRows.Count Then
                notFoundCount = notFoundCount + 1
            End If

        End If

NextFile:

        On Error GoTo ErrorHandler

    Next file

    MsgBox _
        "Inkasso emails have been created as drafts." & vbCrLf & vbCrLf & _
        "Created: " & createdCount & vbCrLf & _
        "Skipped: " & skippedCount & vbCrLf & _
        "Customer not found: " & notFoundCount & vbCrLf & _
        "Errors: " & errorCount, _
        vbInformation, "Inkasso emails"

    Exit Sub


FileError:

    errorCount = errorCount + 1
    Err.Clear
    Resume NextFile


ErrorHandler:

    MsgBox _
        "An unexpected error occurred." & vbCrLf & vbCrLf & _
        "Error: " & Err.Description, _
        vbCritical, "Inkasso Email Automation"

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

    On Error GoTo ErrorHandler

    ' Get Outlook namespace
    Set Ns = OutlookApp.GetNamespace("MAPI")

    ' Get shared mailbox
    Set SharedMailbox = Nothing

    On Error Resume Next
    Set SharedMailbox = Ns.Folders("Zaala Service Center Kunden")
    On Error GoTo ErrorHandler

    If SharedMailbox Is Nothing Then

        MsgBox _
            "Shared mailbox 'Zaala Service Center Kunden' was not found in Outlook.", _
            vbCritical, "Shared mailbox not found"

        MailAlreadySent = False
        Exit Function

    End If

    ' Get Sent Items folder of shared mailbox
    Set SentFolder = Nothing

    On Error Resume Next
    Set SentFolder = SharedMailbox.Folders("Gesendete Elemente")
    On Error GoTo ErrorHandler

    If SentFolder Is Nothing Then

        MsgBox _
            "Folder 'Gesendete Elemente' was not found in shared mailbox.", _
            vbCritical, "Sent Items folder not found"

        MailAlreadySent = False
        Exit Function

    End If

    ' Search only emails sent during the last 30 days
    FilterDate = Date - 30

    Set Items = SentFolder.Items

    ' Sort emails by sent date descending
    Items.Sort "[SentOn]", True

    ' Build Outlook filter
    Filter = "[SentOn] >= '" & _
             Format(FilterDate, "ddddd h:nn AMPM") & "'"

    Set FilteredItems = Items.Restrict(Filter)

    ' Check matching emails
    For Each itm In FilteredItems

        If itm.Class = 43 Then

            If LCase(Trim(itm.Subject)) = _
               LCase(Trim(SubjectText)) Then

                If InStr(1, _
                         LCase(itm.To), _
                         LCase(Recipient), _
                         vbTextCompare) > 0 Then

                    MailAlreadySent = True
                    Exit Function

                End If

            End If

        End If

    Next itm

    MailAlreadySent = False
    Exit Function


ErrorHandler:

    ' If the Sent Items search fails,
    ' allow the email to be created.
    MailAlreadySent = False

End Function



