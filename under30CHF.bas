Attribute VB_Name = "under30CHF"
Option Explicit

'=======================================================================
' Module      : under30CHF
' Procedure   : CreateInkassoEmailUnder30
' Purpose     : Creates Outlook email drafts (per customer) listing
'               unpaid invoices under CHF 30.00 (below the Inkasso
'               transfer limit), based on the newest
'               "*_Inkassoubergabe.xlsx" file found on OneDrive.
'
' Notes       : Business logic is unchanged from the original
'               "CreateInkassoEmailDrafts" procedure. Only the items
'               requested (renaming, subject format, error handling,
'               final statistics) were added.
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

    Dim key As Variant

    ' Statistics counters
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

    filePath = "C:\Users\Desi\OneDrive - Zaala AG\Desislava @ Zaala AG - Linked Files\Inkasso Excel Table\"


    fileName = FindInkassoFile(filePath)


    If fileName = "" Then
        Err.Raise vbObjectError + 1, "CreateInkassoEmailUnder30", _
                  "No Inkasso" & Chr(252) & "bergabe file was found."
    End If


    '-----------------------------------
    ' Extract date from file name
    ' Example: 20260730_Inkassoubergabe.xlsx -> 30.07.2026
    '-----------------------------------

    fileDateRaw = Left(fileName, 8)

    If Not IsDate(Mid(fileDateRaw, 5, 2) & "/" & Right(fileDateRaw, 2) & "/" & Left(fileDateRaw, 4)) Then
        Err.Raise vbObjectError + 2, "CreateInkassoEmailUnder30", _
                  "The date could not be read from the file name: " & fileName
    End If

    fileDateFormatted = Mid(fileDateRaw, 7, 2) & "." & Mid(fileDateRaw, 5, 2) & "." & Left(fileDateRaw, 4)


    '-----------------------------------
    ' Open source file
    '-----------------------------------

    Set wbSource = Workbooks.Open(filePath & fileName, ReadOnly:=True)


    On Error Resume Next
    Set ws = wbSource.Worksheets("<30CHF")
    On Error GoTo ErrorHandler


    If ws Is Nothing Then
        Err.Raise vbObjectError + 3, "CreateInkassoEmailUnder30", _
                  "Worksheet <30CHF was not found."
    End If


    '-----------------------------------
    ' Customer table
    '-----------------------------------

    Set tblCustomers = ThisWorkbook.Worksheets("Kunden").ListObjects("tblKunden")


    Set dictInvoices = CreateObject("Scripting.Dictionary")


    '-----------------------------------
    ' Read invoices
    '-----------------------------------

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row


    For i = 2 To lastRow

        webshopID = Trim(CStr(ws.Cells(i, "A").Value))


        If webshopID <> "" Then

            invoiceLine = _
            "- Rechnung Nr. " & ws.Cells(i, "F").Value & _
            " vom " & ws.Cells(i, "G").Value & _
            ", Rechnungsbetrag CHF " & Format(ws.Cells(i, "H").Value, "0.00") & _
            ", Betrag offen " & Format(ws.Cells(i, "I").Value, "0.00") & _
            ", Patient: " & ws.Cells(i, "J").Value & _
            " " & ws.Cells(i, "K").Value


            If dictInvoices.Exists(webshopID) Then

                dictInvoices(webshopID) = dictInvoices(webshopID) & vbCrLf & invoiceLine

            Else

                dictInvoices.Add webshopID, invoiceLine

            End If

            invoicesRead = invoicesRead + 1

        End If

    Next i


    wbSource.Close False
    Set wbSource = Nothing


    '-----------------------------------
    ' Outlook (late binding)
    '-----------------------------------

    Set OutlookApp = CreateObject("Outlook.Application")


    '-----------------------------------
    ' Create emails
    '-----------------------------------

    For Each key In dictInvoices.Keys

        customerEmail = GetCustomerEmail(CStr(key), tblCustomers)
        language = GetCustomerLanguage(CStr(key), tblCustomers)


        If customerEmail = "" Then

            missingEmailCount = missingEmailCount + 1

        Else

            '-----------------------------------
            ' Subject now includes the Inkassouebergabe date,
            ' so EmailAlreadySent() automatically prevents
            ' duplicate drafts only for the same Inkassouebergabe.
            '-----------------------------------

            If language = "F" Then

                mailSubject = CStr(key) & " | Inkasso" & Chr(252) & "bergabe " & fileDateFormatted & _
                              " | Factures non transmises " & Chr(224) & " recouvrement"

            Else

                mailSubject = CStr(key) & " | Inkasso" & Chr(252) & "bergabe " & fileDateFormatted & _
                              " | " & Chr(220) & "berf" & Chr(228) & "llige Rechnungen"

            End If


            If EmailAlreadySent(OutlookApp, customerEmail, mailSubject) Then

                alreadySentCount = alreadySentCount + 1

            Else

                ' Non-fatal per-item error trapping, so one failed
                ' draft does not abort the whole run.
                On Error Resume Next

                Set mail = OutlookApp.CreateItem(0)

                With mail

                    .SentOnBehalfOfName = "abrechnung@zaala.ch"

                    .To = customerEmail

                    .Subject = mailSubject


                    If language = "F" Then

                        .Body = _
                        "Bonjour," & vbCrLf & vbCrLf & _
                        "Les factures suivantes que vous nous avez soumises n'ont pas " & Chr(233) & "t" & Chr(233) & " pay" & Chr(233) & "es par les patients malgr" & Chr(233) & " plusieurs rappels :" & _
                        vbCrLf & vbCrLf & _
                        dictInvoices(key) & vbCrLf & vbCrLf & _
                        "Comme il s'agit de cr" & Chr(233) & "ances inf" & Chr(233) & "rieures " & Chr(224) & " CHF 30.00 (limite pour le transfert de recouvrement), nous vous prions de nous indiquer si nous devons annuler ces factures ou s'il faut s'attendre " & Chr(224) & " un paiement rapide." & _
                        vbCrLf & vbCrLf & _
                        "Nous vous remercions pour votre attention." & vbCrLf & _
                        "Cordialement," & vbCrLf & _
                        "Centre de service clients" & vbCrLf & vbCrLf & _
                        "Zaala AG" & vbCrLf & _
                        "Bahnhofstrasse 51" & vbCrLf & _
                        "8500 Frauenfeld" & vbCrLf & _
                        "abrechnung@zaala.ch" & vbCrLf & _
                        "www.zaala.ch"

                    Else

                        .Body = _
                        "Guten Tag" & vbCrLf & vbCrLf & _
                        "Folgende Rechnung, die Sie uns eingereicht haben, wurde von dem Patienten trotz mehreren Mahnungen nicht bezahlt:" & _
                        vbCrLf & vbCrLf & _
                        dictInvoices(key) & vbCrLf & vbCrLf & _
                        "Da es sich um eine Forderung unter CHF 30.00 handelt (Limit f" & Chr(252) & "r Inkasso" & Chr(252) & "bergabe), bitten wir Sie, uns mitzuteilen, ob wir diese Rechnungen stornieren sollten oder eine zeitnahe Zahlung zu erwarten ist." & _
                        vbCrLf & vbCrLf & _
                        "Vielen Dank f" & Chr(252) & "r Ihre R" & Chr(252) & "ckmeldung." & vbCrLf & vbCrLf & _
                        "Zaala AG" & vbCrLf & _
                        "Bahnhofstrasse 51" & vbCrLf & _
                        "8500 Frauenfeld" & vbCrLf & _
                        "abrechnung@zaala.ch" & vbCrLf & _
                        "www.zaala.ch"

                    End If


                    .Save

                End With


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
    ' Final statistics (only shown on normal completion)
    '-----------------------------------

    MsgBox "Inkasso Email Under 30 CHF" & vbCrLf & vbCrLf & _
           "Selected file:" & vbCrLf & fileName & vbCrLf & vbCrLf & _
           "Invoices read:" & vbCrLf & invoicesRead & vbCrLf & vbCrLf & _
           "Customers:" & vbCrLf & dictInvoices.Count & vbCrLf & vbCrLf & _
           "Drafts created:" & vbCrLf & draftsCreated & vbCrLf & vbCrLf & _
           "Already sent:" & vbCrLf & alreadySentCount & vbCrLf & vbCrLf & _
           "Missing email:" & vbCrLf & missingEmailCount & vbCrLf & vbCrLf & _
           "Errors:" & vbCrLf & errorsCount, vbInformation


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

    MsgBox "Error " & Err.Number & vbCrLf & Err.Description

    Resume CleanExit

End Sub


'-----------------------------------
' Find newest Inkassoubergabe file
' matches: *_Inkassoubergabe.xlsx
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
' Column F = Merch ID
' Column L = Email
'-----------------------------------

Function GetCustomerEmail(customerID As String, tbl As ListObject) As String

    Dim r As ListRow


    For Each r In tbl.ListRows

        If Trim(CStr(r.Range.Columns(6).Value)) = customerID Then

            GetCustomerEmail = Trim(CStr(r.Range.Columns(12).Value))
            Exit Function

        End If

    Next r


    GetCustomerEmail = ""

End Function



'-----------------------------------
' Get customer language from tblKunden
' Column X
'-----------------------------------

Function GetCustomerLanguage(customerID As String, tbl As ListObject) As String

    Dim r As ListRow
    Dim lang As String


    For Each r In tbl.ListRows

        If Trim(CStr(r.Range.Columns(6).Value)) = customerID Then

            lang = UCase(Trim(CStr(r.Range.Columns(24).Value)))


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

Private Function EmailAlreadySent(ByVal OutlookApp As Object, _
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
    Set SharedMailbox = Ns.Folders("Zaala Service Center Kunden")
    On Error GoTo 0


    If SharedMailbox Is Nothing Then

        EmailAlreadySent = False
        Exit Function

    End If


    Set SentFolder = SharedMailbox.Folders("Gesendete Elemente")


    FilterDate = Date - 30


    Set Items = SentFolder.Items


    Items.Sort "[SentOn]", True


    Filter = "[SentOn] >= '" & Format(FilterDate, "ddddd h:nn AMPM") & "'"


    Set FilteredItems = Items.Restrict(Filter)


    For Each itm In FilteredItems

        If itm.Class = 43 Then

            If LCase(Trim(itm.Subject)) = LCase(Trim(SubjectText)) Then

                If InStr(1, LCase(itm.To), LCase(Recipient), vbTextCompare) > 0 Then

                    EmailAlreadySent = True
                    Exit Function

                End If

            End If

        End If

    Next itm


    EmailAlreadySent = False

End Function
