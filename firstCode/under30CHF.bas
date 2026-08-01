Attribute VB_Name = "under30CHF"
Option Explicit

Sub CreateInkassoEmailDrafts()

    Dim wbSource As Workbook
    Dim ws As Worksheet
    Dim tblCustomers As ListObject
    
    Dim OutlookApp As Object
    Dim mail As Object
    
    Dim filePath As String
    Dim fileName As String
    
    Dim dictInvoices As Object
    
    Dim lastRow As Long
    Dim i As Long
    
    Dim webshopID As String
    Dim customerEmail As String
    Dim language As String
    Dim invoiceLine As String
    Dim mailSubject As String
    
    Dim key As Variant
    
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    
    '-----------------------------------
    ' File location
    '-----------------------------------
    
    filePath = "C:\Users\Desi\OneDrive - Zaala AG\Desislava @ Zaala AG - Linked Files\Inkasso Excel Table\"
    
    
    fileName = FindInkassoFile(filePath)
    
    
    If fileName = "" Then
        
        MsgBox "No Inkassoübergabe file was found.", vbExclamation
        Exit Sub
        
    End If
    
    
    '-----------------------------------
    ' Open source file
    '-----------------------------------
    
    Set wbSource = Workbooks.Open(filePath & fileName, ReadOnly:=True)
    
    
    On Error Resume Next
    Set ws = wbSource.Worksheets("<30CHF")
    On Error GoTo 0
    
    
    If ws Is Nothing Then
        
        MsgBox "Worksheet <30CHF was not found.", vbCritical
        wbSource.Close False
        Exit Sub
        
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
            
        End If
        
    Next i
    
    
    wbSource.Close False
    
    
    '-----------------------------------
    ' Outlook
    '-----------------------------------
    
    Set OutlookApp = CreateObject("Outlook.Application")
    
    
    '-----------------------------------
    ' Create emails
    '-----------------------------------
    
    For Each key In dictInvoices.Keys
        
        
        customerEmail = GetCustomerEmail(CStr(key), tblCustomers)
        language = GetCustomerLanguage(CStr(key), tblCustomers)
        
        
        If customerEmail <> "" Then
            
            
            If language = "F" Then
                
                mailSubject = key & " Factures non transmises à recouvrement."
                
            Else
                
                mailSubject = key & " Überfällige Rechnungen: Beträge unter Inkassolimit."
                
            End If
            
            
            
            If EmailAlreadySent(OutlookApp, customerEmail, mailSubject) = False Then
                
                
                Set mail = OutlookApp.CreateItem(0)
                
                
                With mail
                    
                    .SentOnBehalfOfName = "abrechnung@zaala.ch"
                    
                    .To = customerEmail
                    
                    .Subject = mailSubject
                    
                    
                    If language = "F" Then
                        
                        .Body = _
                        "Bonjour," & vbCrLf & vbCrLf & _
                        "Les factures suivantes que vous nous avez soumises n'ont pas été payées par les patients malgré plusieurs rappels :" & _
                        vbCrLf & vbCrLf & _
                        dictInvoices(key) & vbCrLf & vbCrLf & _
                        "Comme il s'agit de créances inférieures à CHF 30.00 (limite pour le transfert de recouvrement), nous vous prions de nous indiquer si nous devons annuler ces factures ou s’il faut s’attendre à un paiement rapide." & _
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
                        "Da es sich um eine Forderung unter CHF 30.00 handelt (Limit für Inkassoübergabe), bitten wir Sie, uns mitzuteilen, ob wir diese Rechnungen stornieren sollten oder eine zeitnahe Zahlung zu erwarten ist." & _
                        vbCrLf & vbCrLf & _
                        "Vielen Dank für Ihre Rückmeldung." & vbCrLf & vbCrLf & _
                        "Zaala AG" & vbCrLf & _
                        "Bahnhofstrasse 51" & vbCrLf & _
                        "8500 Frauenfeld" & vbCrLf & _
                        "abrechnung@zaala.ch" & vbCrLf & _
                        "www.zaala.ch"
                        
                    End If
                    
                    
                    .Save
                    
                End With
                
                
            End If
            
        End If
        
    Next key
    
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    
    MsgBox "Inkasso email drafts have been created.", vbInformation


End Sub

'-----------------------------------
' Find newest Inkassoubergabe file
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

        If InStr(1, LCase(f.Name), "_inkasso", vbTextCompare) > 0 _
           And LCase(Right(f.Name, 5)) = ".xlsx" Then

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
