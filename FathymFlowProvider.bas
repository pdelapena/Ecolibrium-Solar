Attribute VB_Name = "FathymFlowProvider"
Public Const FathymFlowUrl As String = "https://flw-ecosol-prd.azurewebsites.net/api/HttpAlertToSugarCRM?code=2c907115-e1a9-485e-b5a3-cf7f79fcf3bc"
Public SugarAccount As String
Public SugarAccountId As String
Public SugarOpportunityId As String
Public SugarContactId As String
Public SugarQuoteId As String
Public ExistingRecord As Boolean
Public ActiveSheet As String

Sub SendToSugarCRM()
    Application.ScreenUpdating = False
    Sheets("IN-N-OUT").Visible = True
    
    ActiveSheet = Application.ActiveSheet.Name
    Sheets("IN-N-OUT").Select
    
    ExcelProvider.Initialize

    ApplicationProvider.CallsMade = 0
    ApplicationProvider.TotalCalls = 10 + ExcelProvider.LineItemCount
    
    SugarAccountId = ExcelProvider.GetCellValue("Accounts", "id")
    
    If (SugarAccountId <> vbNullString And SugarAccountId <> "0") Then
        SendDataToFathymFlow
    Else
        frmProgress.Hide
        
        Dim accountName As String
        
        accountName = JsonConverter.ParseJson(ExcelProvider.GetAccountName(vbNullString))("name")
        
        frmStartup.lblAccountName.Caption = accountName
        
        frmStartup.Show (0)
    End If
End Sub

Public Function SendDataToFathymFlow() As String
    On Error GoTo HandleError
    Application.ScreenUpdating = False
    
    Dim StartTime As Double
    Dim MinutesElapsed As String
    
    StartTime = Timer
    
    frmProgress.Show (0)
    
    ApplicationProvider.SetStatus ("Sending of Excel Data to Fathym Flow Started..")
    
    Dim linkResp As Boolean
    Dim lineItems As Collection
    
    If (ExistingRecord = False) Then
        ApplicationProvider.TotalCalls = ApplicationProvider.TotalCalls + 5
    End If
    
    ApplicationProvider.UpdateProgress
    
    If (SugarAccountId = vbNullString Or SugarAccountId = "0") Then
        SugarAccount = CheckDuplicateAccount
        ApplicationProvider.UpdateProgress
        SugarAccountId = CreateOrUpdateAccount()
    End If
        
    If (ExistingRecord = False) Then
        ExcelProvider.SetModuleId "Accounts", SugarAccountId
    End If
    
    ApplicationProvider.UpdateProgress
    Dim contact As String
    If (SugarContactId = vbNullString Or SugarContactId = "0") Then
        contact = CheckDuplicateContact
    End If
    
    ApplicationProvider.UpdateProgress
    SugarContactId = CreateOrUpdateContact()
    
    If (ExistingRecord = False) Then
        ApplicationProvider.UpdateProgress
        ExcelProvider.SetModuleId "Contacts", SugarContactId
        linkResp = LinkAccountContact(SugarAccountId, SugarContactId)
    End If
    
    ApplicationProvider.UpdateProgress
    SugarOpportunityId = CreateOrUpdateOpportunity(SugarAccountId)
    
    If (ExistingRecord = False) Then
        ApplicationProvider.UpdateProgress
        ExcelProvider.SetModuleId "Opportunities", SugarOpportunityId
        linkResp = LinkAccountOpportunity(SugarAccountId, SugarOpportunityId)
    End If
         
    If (ExistingRecord = False) Then
        ApplicationProvider.UpdateProgress
        SugarQuoteId = CreateOrUpdateQuote()
        ExcelProvider.SetModuleId "Quotes", SugarQuoteId
    End If
    
    
    If (ExistingRecord = False) Then
        ApplicationProvider.UpdateProgress
        'WILL UPDATE BOM2
        'linkResp = LinkQuoteOpportunity(SugarOpportunityId, SugarQuoteId)
        
        ApplicationProvider.UpdateProgress
        linkResp = LinkQuoteAccount(SugarAccountId, SugarQuoteId)
    End If
    
    
    Set lineItems = CreateOrUpdateQuotedLineItems(SugarQuoteId)
    SetLineItemIds lineItems
    
    
    ApplicationProvider.UpdateProgress
    SendDataToFathymFlow = SugarAccountId
    
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    SetStatus ("This code ran successfully in " & MinutesElapsed & " minutes.")
    
    Sheets(ActiveSheet).Select
    Sheets("IN-N-OUT").Visible = False
    
    Application.ScreenUpdating = True
    
    Application.ActiveWorkbook.Save
    
    ApplicationProvider.SetComplete ("Successfully synced Excel doc with SugarCRM.")
HandleError:
    If Err.Number <> 0 Then
        ApplicationProvider.SetError
        
        frmProgress.Hide
        Application.ScreenUpdating = True
        Application.ActiveWorkbook.Save
        
        Sheets(ActiveSheet).Select
        Sheets("IN-N-OUT").Visible = False
    End If
End Function

Public Function CheckDuplicateAccount() As String
    
    ApplicationProvider.SetStatus ("Checking Duplicate Accounts..")
    
    Dim request As New Dictionary
        
    request("ActionType") = "DuplicateCheck"
    request("ModuleType") = "Accounts"
    request("Object") = ExcelProvider.GetAccountName(SugarAccount)
 
    SugarAccount = JsonConverter.ParseJson(request("Object"))("name")
 
    Dim response As Object
    
    Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
    
    If (response.Status <> 200) Then
        ApplicationProvider.SetStatus ("Found Duplicate Accounts..")
        
        Dim responseObject As New Dictionary
        Dim dupes As New Collection
        
        Set responseObject = JsonConverter.ParseJson(response.ResponseText)
        
        Set dupes = responseObject("records")
        
        frmDuplicates.cmbAccounts.Clear
        
        If (dupes.count > 0) Then
            For i = 1 To dupes.count
                Dim duplicate As Object
                Set duplicate = dupes(i)
                frmDuplicates.cmbAccounts.AddItem duplicate("name")
                frmDuplicates.cmbAccountIds.AddItem duplicate("id")
            Next i
            
            frmProgress.Hide
            frmDuplicates.Show (0)
            Do While frmDuplicates.Visible
                DoEvents
            Loop
        End If
              
        ApplicationProvider.SetStatus ("Sugar Account: " & SugarAccount)
    End If
    
    CheckDuplicateAccount = SugarAccount
    
End Function

Public Function CheckDuplicateContact() As String
    
    ApplicationProvider.SetStatus ("Checking Duplicate Contacts..")
    
    Dim request As New Dictionary
        
    request("ActionType") = "DuplicateCheck"
    request("ModuleType") = "Contacts"
    request("Object") = ExcelProvider.GetContact(SugarContactId)
 
    Dim sugarContact As String
    sugarContact = JsonConverter.ParseJson(request("Object"))("name")
 
    Dim response As Object
    
    Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
    
    If (response.Status <> 200) Then
        ApplicationProvider.SetStatus ("Found Duplicate Accounts..")
        
        Dim responseObject As New Dictionary
        Dim dupes As New Collection
        
        Set responseObject = JsonConverter.ParseJson(response.ResponseText)
        
        Set dupes = responseObject("records")
        If (dupes.count > 0) Then
            For i = 1 To dupes.count
                Dim duplicate As Object
                Set duplicate = dupes(i)
                If (duplicate("account_id") = SugarAccountId Or duplicate("account_name") = SugarAccount) Then
                    SugarContactId = duplicate("id")
                    Exit For
                End If
            Next i
        End If
               
        ApplicationProvider.SetStatus ("Sugar Contact: " & SugarContactId)
    End If
    
    CheckDuplicateContact = sugarContact
    
End Function

Public Function CreateOrUpdateAccount() As String
    
    ApplicationProvider.SetStatus ("Creating/Updating Account: " & SugarAccount)

    Dim request As New Dictionary
    
    request("ActionType") = "CreateOrUpdate"
    request("ModuleType") = "Accounts"
    request("Object") = ExcelProvider.GetAccount(SugarAccount, SugarAccountId)

    Dim response As Object
    
    Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
    
    If (response.Status <> 200) Then
        ApplicationProvider.SetError (response.ResponseText)
    End If
    
    Dim responseObject As New Dictionary
    Set responseObject = JsonConverter.ParseJson(response.ResponseText)
    
    ApplicationProvider.SetStatus ("Done Creating/Updating Account: " & SugarAccount & " : " & responseObject("id"))
    
    CreateOrUpdateAccount = responseObject("id")
    
End Function

Public Function CreateOrUpdateContact() As String
    
    ApplicationProvider.SetStatus ("Creating/Updating Contact")

    Dim request As New Dictionary
    
    request("ActionType") = "CreateOrUpdate"
    request("ModuleType") = "Contacts"
    request("Object") = ExcelProvider.GetContact(SugarContactId)

    Dim response As Object
    
    Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
    
    If (response.Status <> 200) Then
        ApplicationProvider.SetError (response.ResponseText)
    End If
    
    Dim responseObject As New Dictionary
    Set responseObject = JsonConverter.ParseJson(response.ResponseText)
    
    ApplicationProvider.SetStatus ("Done Creating/Updating Contact: " & responseObject("id"))
    
    CreateOrUpdateContact = responseObject("id")
    
End Function

Public Function CreateOrUpdateOpportunity(ByVal accountId As String) As String
    
    ApplicationProvider.SetStatus ("Creating/Updating Opportunity for Account: " & accountId)

    Dim request As New Dictionary
    
    request("ActionType") = "CreateOrUpdate"
    request("ModuleType") = "Opportunities"
    request("Object") = ExcelProvider.GetOpportunity(SugarOpportunityId)
    
    Dim response As Object
    
    Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
    
    If (response.Status <> 200) Then
        ApplicationProvider.SetError (response.ResponseText)
    End If
    
    Dim responseObject As New Dictionary
    Set responseObject = JsonConverter.ParseJson(response.ResponseText)
    
    ApplicationProvider.SetStatus ("Done Creating/Updating Opportunity: " & responseObject("id"))

    CreateOrUpdateOpportunity = responseObject("id")
    
End Function

Public Function CreateOrUpdateQuote() As String
    
    ApplicationProvider.SetStatus ("Creating/Updating Quote")

    Dim request As New Dictionary
    
    request("ActionType") = "CreateOrUpdate"
    request("ModuleType") = "Quotes"
    request("Object") = ExcelProvider.GetQuote(SugarQuoteId, SugarOpportunityId)

    Dim response As Object
    
    Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
    
    If (response.Status <> 200) Then
        ApplicationProvider.SetError (response.ResponseText)
    End If
    
    Dim responseObject As New Dictionary
    Set responseObject = JsonConverter.ParseJson(response.ResponseText)
    
    ApplicationProvider.SetStatus ("Done Creating/Updating Quote: " & responseObject("id"))
    
    CreateOrUpdateQuote = responseObject("id")
    
End Function

Public Function CreateOrUpdateQuotedLineItems(ByVal quoteId As String) As Variant
    ApplicationProvider.SetStatus ("Creating/Updating Quoted Line Items")
    
    Dim request As New Dictionary
    Dim products As New Collection
 
    Dim lineItem As Variant
    
    ApplicationProvider.UpdateProgress
    For Each lineItem In ExcelProvider.GetQuotedLineItems(quoteId, SugarAccountId, SugarOpportunityId)
        
    
        Dim itm As String
        itm = JsonConverter.ConvertToJson(lineItem, Whitespace:=3)
        
        ApplicationProvider.SetStatus ("Creating Line Item: " & lineItem("name"))
    
        request("ModuleType") = "Products"
        request("Object") = itm
    
        Dim response As Object
        
        If (lineItem("id") <> vbNullString And lineItem("quantity") = 0) Then
            request("ActionType") = "Delete"
        Else
            request("ActionType") = "CreateOrUpdate"
        End If
        
        Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
        
        ApplicationProvider.UpdateProgress
        
        If (response.Status <> 200) Then
            ApplicationProvider.SetError (response.ResponseText)
        End If
        
        Dim responseObject As New Dictionary
        Set responseObject = JsonConverter.ParseJson(response.ResponseText)
        
        SetStatus ("CreateOrUpdateQuotedLineItems {" & lineItem("name") & "}")
        products.add responseObject
    Next lineItem
    
    ApplicationProvider.UpdateProgress
    Dim resp As Boolean
    resp = UpdateQuote(quoteId, products)
 
    Set CreateOrUpdateQuotedLineItems = products
  
    ApplicationProvider.SetStatus ("Done Creating/Updating Quoted Line Items")
End Function

Public Function LinkAccountContact(ByVal accountId As String, ByVal contactId As String) As Boolean
    
    ApplicationProvider.SetStatus ("Linking Account: " & accountId & " To Contact: " & contactId)

    Dim requestObject As New Dictionary
    Dim request As New Dictionary
    
    requestObject("modulePath") = "Contacts"
    requestObject("moduleId") = contactId
    requestObject("linkedModulePath") = "Accounts"
    requestObject("linkedModuleId") = accountId
    
    request("ActionType") = "Link"
    request("ModuleType") = "Contacts"
    request("Object") = JsonConverter.ConvertToJson(requestObject, Whitespace:=3)
    
    Dim response As Object
    
    Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
    
    If (response.Status <> 200) Then
        ApplicationProvider.SetError (response.ResponseText)
    End If
    
    Dim responseObject As New Dictionary
    Set responseObject = JsonConverter.ParseJson(response.ResponseText)
    
    ApplicationProvider.SetStatus ("Done Linking Account: " & accountId & " To Contact: " & contactId)
    
    LinkAccountContact = True
    
End Function

Public Function LinkAccountOpportunity(ByVal accountId As String, ByVal opportunityId As String) As Boolean
    
    ApplicationProvider.SetStatus ("Linking Account: " & accountId & " To Opportunity: " & opportunityId)

    Dim requestObject As New Dictionary
    Dim request As New Dictionary
    
    requestObject("modulePath") = "Opportunities"
    requestObject("moduleId") = opportunityId
    requestObject("linkedModulePath") = "Accounts"
    requestObject("linkedModuleId") = accountId
    
    request("ActionType") = "Link"
    request("ModuleType") = "Opportunities"
    request("Object") = JsonConverter.ConvertToJson(requestObject, Whitespace:=3)
    
    Dim response As Object
    
    Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
    
    If (response.Status <> 200) Then
        ApplicationProvider.SetError (response.ResponseText)
    End If
    
    Dim responseObject As New Dictionary
    Set responseObject = JsonConverter.ParseJson(response.ResponseText)
    
    ApplicationProvider.SetStatus ("Done Linking Account: " & accountId & " To Opportunity: " & opportunityId)
    
    LinkAccountOpportunity = True
    
End Function

Public Function LinkQuoteOpportunity(ByVal opportunityId As String, ByVal quoteId As String) As Boolean
    
    ApplicationProvider.SetStatus ("Linking Quote: " & quoteId & " To Opportunity: " & opportunityId)

    Dim requestObject As New Dictionary
    Dim request As New Dictionary
    
    requestObject("modulePath") = "Opportunities"
    requestObject("moduleId") = opportunityId
    requestObject("linkedModulePath") = "Quotes"
    requestObject("linkedModuleId") = quoteId
    
    request("ActionType") = "Link"
    request("ModuleType") = "Quotes"
    request("Object") = JsonConverter.ConvertToJson(requestObject, Whitespace:=3)
    
    Dim response As Object
    
    Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
    
    If (response.Status <> 200) Then
        ApplicationProvider.SetError (response.ResponseText)
    End If
    
    Dim responseObject As New Dictionary
    Set responseObject = JsonConverter.ParseJson(response.ResponseText)
    
    ApplicationProvider.SetStatus ("Done Linking Quote: " & quoteId & " To Opportunity: " & opportunityId)
    
    LinkQuoteOpportunity = True
    
End Function

Public Function LinkQuoteAccount(ByVal accountId As String, ByVal quoteId As String) As Boolean
    
    ApplicationProvider.SetStatus ("Linking Quote: " & quoteId & " To Account: " & accountId)

    Dim requestObject As New Dictionary
    Dim request As New Dictionary
    
    requestObject("modulePath") = "Accounts"
    requestObject("moduleId") = accountId
    requestObject("linkedModulePath") = "Quotes"
    requestObject("linkedModuleId") = quoteId
    
    request("ActionType") = "Link"
    request("ModuleType") = "Quotes"
    request("Object") = JsonConverter.ConvertToJson(requestObject, Whitespace:=3)
    
    Dim response As Object
    
    Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
    
    If (response.Status <> 200) Then
        ApplicationProvider.SetError (response.ResponseText)
    End If
    
    Dim responseObject As New Dictionary
    Set responseObject = JsonConverter.ParseJson(response.ResponseText)
    
    ApplicationProvider.SetStatus ("Done Linking Quote: " & quoteId & " To Account: " & accountId)
    
    LinkQuoteAccount = True
    
End Function

Public Function LookupAccounts(ByVal lookup As String) As Collection
    
    Dim request As New Dictionary
    
    Dim account As New Dictionary
    account("name") = lookup
    
    request("ActionType") = "DuplicateCheck"
    request("ModuleType") = "Accounts"
    request("Object") = JsonConverter.ConvertToJson(account, Whitespace:=3)
 
    Dim response As Object
    
    Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
    
    Dim dupes As New Collection
    If (response.Status <> 200) Then
        Dim responseObject As New Dictionary
        
        Set responseObject = JsonConverter.ParseJson(response.ResponseText)
        
        Set dupes = responseObject("records")
    End If
    
    Set LookupAccounts = dupes
    
End Function

Public Function SetLineItemIds(ByVal lineItems As Collection) As Boolean
    
    Dim itm As Variant
    
    For Each itm In lineItems
        If (itm("quantity") = 0 And itm("id") <> vbNullString) Then
            ExcelProvider.SetLineItemId itm("name"), ""
        Else
            ExcelProvider.SetLineItemId itm("name"), itm("id")
        End If
    Next itm
    
End Function

Public Function UpdateQuote(ByVal quoteId As String, ByVal products As Collection) As Boolean
    
    If (products.count > 0) Then
        ApplicationProvider.SetStatus ("Updating Quote " & quoteId & " With " & products.count & " QuotedLineItems")
    
        Dim request As New Dictionary
        
        request("ActionType") = "CreateOrUpdate"
        request("ModuleType") = "Quotes"
        request("Object") = ExcelProvider.GetQuote(quoteId, SugarOpportunityId, products)
    
        Dim response As Object
        
        Set response = HttpPost(FathymFlowUrl, JsonConverter.ConvertToJson(request, Whitespace:=3))
        
        If (response.Status <> 200) Then
            ApplicationProvider.SetError (response.ResponseText)
        End If
        
        Dim responseObject As New Dictionary
        Set responseObject = JsonConverter.ParseJson(response.ResponseText)
        
        ApplicationProvider.SetStatus ("Updated Quote " & quoteId & " QuotedLineItems")
    End If
    
    UpdateQuote = True
    
End Function
