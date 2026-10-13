Attribute VB_Name = "under30CHF"
Option Explicit

'=======================================================================
' Module      : under30CHF
' Procedure   : CreateInkassoEmailUnder30
' Purpose     : Creates Outlook email drafts for unpaid invoices
'               under CHF 30.00.
'
' tblKunden:
' E = Merch ID
' G = eMail für Dokumente
' H = eMail Praxismanager (fallback)
' W = Sprache
'=======================================================================

Sub CreateInkassoEmailUnder30()

    On Error GoTo ErrorHandler

    Dim wbSource As Workbook
    Dim ws As Worksheet
    Dim tblCustomers As ListObject

    Dim OutlookApp As Object
    Dim mail As Object

    Dim filePath As String
    Dim fileName As String
    Dim fileDateRaw As String
    Dim fileDateFormatted As String

    Dim dictInvoices As Object

    Dim lastRow As Long
    Dim i As Long

    Dim webshopID As String
    Dim customerEmail As String
    Dim language As String
    Dim invoiceLine As String
    Dim mailSubject As String
    Dim mailBody As String

    Dim key As Variant

    Dim invoicesRead As Long
    Dim draftsCreated As Long
    Dim alreadySentCount As Long
    Dim missingEmailCount As Long
    Dim errorsCount As Long

    invoicesRead = 0
    draftsCreated = 0
    alreadySentCount = 0
    missingEmailCount = 0
    errorsCount = 0

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '-----------------------------------
    ' File location
    '-----------------------------------

    filePath = "xxx\Inkasso Excel Table\"

    fileName = FindInkassoFile(filePath)

    If fileName = "" Then
        Err.Raise vbObjectError + 1, _
                  "CreateInkassoEmailUnder30", _
                  "No Inkassoubergabe file was found."
    End If

    '-----------------------------------
    ' Extract date from file name
    ' Example: 20260730_Inkassoubergabe.xlsx
    '-----------------------------------

    fileDateRaw = Left(fileName, 8)

    If Not IsDate(Mid(fileDateRaw, 5, 2) & "/" & _
                  Right(fileDateRaw, 2) & "/" & _
                  Left(fileDateRaw, 4)) Then

        Err.Raise vbObjectError + 2, _
                  "CreateInkassoEmailUnder30", _
                  "The date could not be read from the file name: " & fileName
    End If

    fileDateFormatted = _
        Mid(fileDateRaw, 7, 2) & "." & _
        Mid(fileDateRaw, 5, 2) & "." & _
        Left(fileDateRaw, 4)

    '-----------------------------------
    ' Open source file
    '-----------------------------------

    Set wbSource = Workbooks.Open( _
        filePath & fileName, _
        ReadOnly:=True)

    On Error Resume Next
    Set ws = wbSource.Worksheets("<30CHF")
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Err.Raise vbObjectError + 3, _
                  "CreateInkassoEmailUnder30", _
                  "Worksheet <30CHF was not found."
    End If

    '-----------------------------------
    ' Customer table
    '-----------------------------------

    Set tblCustomers = _
        ThisWorkbook.Worksheets("Kunden").ListObjects("tblKunden")

    Set dictInvoices = CreateObject("Scripting.Dictionary")

    '-----------------------------------
    ' Read invoices
    '-----------------------------------

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow

        webshopID = Trim(CStr(ws.Cells(i, "A").Value))

        If webshopID <> "" Then

            invoiceLine = "- Rechnung Nr. " & ws.Cells(i, "F").Value
            invoiceLine = invoiceLine & " vom " & ws.Cells(i, "G").Value
            invoiceLine = invoiceLine & ", Rechnungsbetrag CHF "
            invoiceLine = invoiceLine & Format(ws.Cells(i, "H").Value, "0.00")
            invoiceLine = invoiceLine & ", Betrag offen "
            invoiceLine = invoiceLine & Format(ws.Cells(i, "I").Value, "0.00")
            invoiceLine = invoiceLine & ", Patient: "
            invoiceLine = invoiceLine & ws.Cells(i, "J").Value
            invoiceLine = invoiceLine & " " & ws.Cells(i, "K").Value

            If dictInvoices.Exists(webshopID) Then

                dictInvoices(webshopID) = _
                    dictInvoices(webshopID) & _
                    vbCrLf & invoiceLine

            Else

                dictInvoices.Add webshopID, invoiceLine

            End If

            invoicesRead = invoicesRead + 1

        End If

    Next i

    wbSource.Close False
    Set wbSource = Nothing

    '-----------------------------------
    ' Outlook
    '-----------------------------------

    Set OutlookApp = CreateObject("Outlook.Application")

    '-----------------------------------
    ' Create emails
    '-----------------------------------

    For Each key In dictInvoices.Keys

        customerEmail = GetCustomerEmail( _
            CStr(key), _
            tblCustomers)

        language = GetCustomerLanguage( _
            CStr(key), _
            tblCustomers)

        If customerEmail = "" Then

            missingEmailCount = missingEmailCount + 1

        Else

            '-----------------------------------
            ' Subject
            '-----------------------------------

            If language = "F" Then

                mailSubject = CStr(key)
                mailSubject = mailSubject & " | Inkasso" & ChrW(252)
                mailSubject = mailSubject & "bergabe " & fileDateFormatted
                mailSubject = mailSubject & " | Factures non transmises "
                mailSubject = mailSubject & ChrW(224)
                mailSubject = mailSubject & " recouvrement"

            Else

                mailSubject = CStr(key)
                mailSubject = mailSubject & " | Inkasso" & ChrW(252)
                mailSubject = mailSubject & "bergabe " & fileDateFormatted
                mailSubject = mailSubject & " | " & ChrW(220)
                mailSubject = mailSubject & "berf" & ChrW(228)
                mailSubject = mailSubject & "llige Rechnungen"

            End If

            '-----------------------------------
            ' Check Sent Items
            '-----------------------------------

            If EmailAlreadySent( _
                OutlookApp, _
                customerEmail, _
                mailSubject) Then

                alreadySentCount = alreadySentCount + 1

            Else

                '-----------------------------------
                ' Create draft
                '-----------------------------------

                On Error Resume Next

                Set mail = OutlookApp.CreateItem(0)

                If language = "F" Then

                    mailBody = "Bonjour,"
                    mailBody = mailBody & vbCrLf & vbCrLf

                    mailBody = mailBody & _
                        "Les factures suivantes que vous nous avez soumises n'ont pas "
                    mailBody = mailBody & "été payées par les patients malgré plusieurs rappels :"
                    mailBody = mailBody & vbCrLf & vbCrLf

                    mailBody = mailBody & dictInvoices(key)
                    mailBody = mailBody & vbCrLf & vbCrLf

                    mailBody = mailBody & _
                        "Comme il s'agit de créances inférieures à CHF 30.00 "
                    mailBody = mailBody & _
                        "(limite pour le transfert de recouvrement), "
                    mailBody = mailBody & _
                        "nous vous prions de nous indiquer si nous devons "
                    mailBody = mailBody & _
                        "annuler ces factures ou s'il faut s'attendre à un paiement rapide."

                    mailBody = mailBody & vbCrLf & vbCrLf
                    mailBody = mailBody & _
                        "Nous vous remercions pour votre attention."
                    mailBody = mailBody & vbCrLf
                    mailBody = mailBody & "Cordialement,"
                    mailBody = mailBody & vbCrLf
                    mailBody = mailBody & "Centre de service clients"
                    mailBody = mailBody & vbCrLf & vbCrLf
                    mailBody = mailBody & "Zaala AG"
                    mailBody = mailBody & vbCrLf
                    mailBody = mailBody & "Bahnhofstrasse 51"
                    mailBody = mailBody & vbCrLf
                    mailBody = mailBody & "8500 Frauenfeld"
                    mailBody = mailBody & vbCrLf
                    mailBody = mailBody & "xxx@zaala.ch"
                    mailBody = mailBody & vbCrLf
                    mailBody = mailBody & "www.zaala.ch"

                Else

                    mailBody = "Guten Tag"
                    mailBody = mailBody & vbCrLf & vbCrLf

                    mailBody = mailBody & _
                        "Folgende Rechnung, die Sie uns eingereicht haben, "
                    mailBody = mailBody & _
                        "wurde von dem Patienten trotz mehreren Mahnungen "
                    mailBody = mailBody & "nicht bezahlt:"
                    mailBody = mailBody & vbCrLf & vbCrLf

                    mailBody = mailBody & dictInvoices(key)
                    mailBody = mailBody & vbCrLf & vbCrLf

                    mailBody = mailBody & _
                        "Da es sich um eine Forderung unter CHF 30.00 handelt "
                    mailBody = mailBody & _
                        "(Limit für Inkassoübergabe), "
                    mailBody = mailBody & _
                        "bitten wir Sie, uns mitzuteilen, ob wir diese "
                    mailBody = mailBody & _
                        "Rechnungen stornieren sollten oder eine zeitnahe "
                    mailBody = mailBody & _
                        "Zahlung zu erwarten ist."

                    mailBody = mailBody & vbCrLf & vbCrLf

                    mailBody = mailBody & _
                        "Vielen Dank für Ihre Rückmeldung."
                    mailBody = mailBody & vbCrLf & vbCrLf
                    mailBody = mailBody & "Zaala AG"
                    mailBody = mailBody & vbCrLf
                    mailBody = mailBody & "Bahnhofstrasse 51"
                    mailBody = mailBody & vbCrLf
                    mailBody = mailBody & "8500 Frauenfeld"
                    mailBody = mailBody & vbCrLf
                    mailBody = mailBody & "xxx@zaala.ch"
                    mailBody = mailBody & vbCrLf
                    mailBody = mailBody & "www.zaala.ch"

                End If

                If Err.Number = 0 Then

                    With mail

                        .SentOnBehalfOfName = "xxx@zaala.ch"
                        .To = customerEmail
                        .Subject = mailSubject
                        .Body = mailBody
                        .Save

                    End With

                End If

                If Err.Number <> 0 Then

                    errorsCount = errorsCount + 1
                    Err.Clear

                Else

                    draftsCreated = draftsCreated + 1

                End If

                Set mail = Nothing

                On Error GoTo ErrorHandler

            End If

        End If

    Next key

    '-----------------------------------
    ' Final statistics
    '-----------------------------------

    MsgBox _
        "Inkasso Email Under 30 CHF" & vbCrLf & vbCrLf & _
        "Selected file:" & vbCrLf & fileName & vbCrLf & vbCrLf & _
        "Invoices read:" & vbCrLf & invoicesRead & vbCrLf & vbCrLf & _
        "Customers:" & vbCrLf & dictInvoices.Count & vbCrLf & vbCrLf & _
        "Drafts created:" & vbCrLf & draftsCreated & vbCrLf & vbCrLf & _
        "Already sent:" & vbCrLf & alreadySentCount & vbCrLf & vbCrLf & _
        "Missing email:" & vbCrLf & missingEmailCount & vbCrLf & vbCrLf & _
        "Errors:" & vbCrLf & errorsCount, _
        vbInformation, _
        "Inkasso Email Under 30 CHF"

CleanExit:

    If Not wbSource Is Nothing Then

        On Error Resume Next
        wbSource.Close False
        On Error GoTo 0

        Set wbSource = Nothing

    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    Set mail = Nothing
    Set OutlookApp = Nothing
    Set dictInvoices = Nothing
    Set tblCustomers = Nothing
    Set ws = Nothing

    Exit Sub

ErrorHandler:

    MsgBox _
        "Error " & Err.Number & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Inkasso Email Under 30 CHF"

    Resume CleanExit

End Sub


'-----------------------------------
' Find newest Inkassoubergabe file
' Matches: *_Inkassoubergabe.xlsx
'-----------------------------------

Function FindInkassoFile(folderPath As String) As String

    Dim fso As Object
    Dim folder As Object
    Dim f As Object

    Dim newestFile As String
    Dim newestDate As Date

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    newestDate = #1/1/1900#
    newestFile = ""

    For Each f In folder.Files

        If LCase(f.Name) Like "*_inkassoubergabe.xlsx" _
           And Left(f.Name, 2) <> "~$" Then

            If f.DateLastModified > newestDate Then

                newestDate = f.DateLastModified
                newestFile = f.Name

            End If

        End If

    Next f

    FindInkassoFile = newestFile

End Function


'-----------------------------------
' Get customer email from tblKunden
'
' Column E = Merch ID
' Column G = eMail für Dokumente
' Column H = eMail Praxismanager (fallback)
'-----------------------------------

Function GetCustomerEmail( _
    customerID As String, _
    tbl As ListObject) As String

    Dim r As ListRow
    Dim emailDocuments As String
    Dim emailPraxismanager As String

    For Each r In tbl.ListRows

        If Trim(CStr(r.Range.Columns(5).Value)) = customerID Then

            emailDocuments = _
                Trim(CStr(r.Range.Columns(7).Value))

            emailPraxismanager = _
                Trim(CStr(r.Range.Columns(8).Value))

            If emailDocuments <> "" Then

                GetCustomerEmail = emailDocuments

            Else

                GetCustomerEmail = emailPraxismanager

            End If

            Exit Function

        End If

    Next r

    GetCustomerEmail = ""

End Function


'-----------------------------------
' Get customer language from tblKunden
'
' Column E = Merch ID
' Column W = Sprache
'-----------------------------------

Function GetCustomerLanguage( _
    customerID As String, _
    tbl As ListObject) As String

    Dim r As ListRow
    Dim lang As String

    For Each r In tbl.ListRows

        If Trim(CStr(r.Range.Columns(5).Value)) = customerID Then

            lang = UCase( _
                Trim(CStr(r.Range.Columns(23).Value)))

            If Left(lang, 1) = "F" Then

                GetCustomerLanguage = "F"

            Else

                GetCustomerLanguage = "D"

            End If

            Exit Function

        End If

    Next r

    GetCustomerLanguage = "D"

End Function


'-----------------------------------
' Check Sent Items
'-----------------------------------

Private Function EmailAlreadySent( _
    ByVal OutlookApp As Object, _
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

    Set Ns = OutlookApp.GetNamespace("MAPI")

    On Error Resume Next
    Set SharedMailbox = _
        Ns.Folders("Zaala Service Center Kunden")
    On Error GoTo 0

    If SharedMailbox Is Nothing Then

        EmailAlreadySent = False
        Exit Function

    End If

    On Error Resume Next
    Set SentFolder = _
        SharedMailbox.Folders("Gesendete Elemente")
    On Error GoTo 0

    If SentFolder Is Nothing Then

        EmailAlreadySent = False
        Exit Function

    End If

    FilterDate = Date - 30

    Set Items = SentFolder.Items

    Items.Sort "[SentOn]", True

    Filter = _
        "[SentOn] >= '" & _
        Format(FilterDate, "ddddd h:nn AMPM") & "'"

    Set FilteredItems = Items.Restrict(Filter)

    For Each itm In FilteredItems

        If itm.Class = 43 Then

            If LCase(Trim(itm.Subject)) = _
               LCase(Trim(SubjectText)) Then

                If InStr( _
                    1, _
                    LCase(itm.To), _
                    LCase(Recipient), _
                    vbTextCompare) > 0 Then

                    EmailAlreadySent = True
                    Exit Function

                End If

            End If

        End If

    Next itm

    EmailAlreadySent = False

End Function