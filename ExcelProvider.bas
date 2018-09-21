Attribute VB_Name = "ExcelProvider"
Public AccountRows As New Collection
Public ContactRows As New Collection
Public OpportunityRows As New Collection
Public QuoteRows As New Collection
Public QuotedLineItemRows As New Collection
Public LineItemCount As Integer

Public Function Initialize()
    SetVariables
    
    Set QuotedLineItemRows = GetLineItemRowNumbers()
End Function

Public Function GetAccount(ByVal accountName As String, ByVal accountId As String) As String

    Dim account As New Dictionary
    Dim rowNum As Variant
    
    For Each rowNum In AccountRows
        Dim varName As String
        Dim varVal As String
        
        varName = CStr(Cells(CInt(rowNum), 6))
        varVal = CStr(Cells(CInt(rowNum), 3))
        
        If (varName <> "" And varName <> "id" And varName <> "name") Then
            account(varName) = varVal
        End If
    Next rowNum
    
    Set account("email") = GetEmails
    
    If (accountId <> vbNullString And accountId <> "0") Then
        account("id") = accountId
    End If
    
    If (accountName <> vbNullString And accountName <> "0") Then
        account("name") = accountName
    End If
    
    GetAccount = JsonConverter.ConvertToJson(account, Whitespace:=3)
    
End Function

Public Function GetContact(ByVal contactId As String) As String

    Dim contact As New Dictionary
    Dim rowNum As Variant
    
    For Each rowNum In ContactRows
        Dim varName As String
        Dim varVal As String
        varName = CStr(Cells(CInt(rowNum), 6))
        varVal = CStr(Cells(CInt(rowNum), 3))
        If (varName <> "" And varName <> "id") Then
            contact(varName) = varVal
        End If
    Next rowNum
    
    Dim nameParts() As String
    Dim firstName As String
    Dim lastName As String
    nameParts = VBA.Split(GetCellValue("Contacts", "name"), " ")
    
    Dim element As Variant

    Dim first As Boolean
    
    For Each element In nameParts
        If (first = False) Then
            firstName = CStr(element)
            first = True
        Else
            If (lastName = vbNullString) Then
                lastName = CStr(element)
            Else
                lastName = lastName & " " & CStr(element)
            End If
        End If
    Next element
    
    contact("first_name") = firstName
    contact("last_name") = lastName
    
    Set contact("email") = GetEmails
    
    If (contactId <> vbNullString And contactId <> "0") Then
        contact("id") = contactId
    End If
    
    GetContact = JsonConverter.ConvertToJson(contact, Whitespace:=3)
    
End Function

Public Function GetEmails() As Collection

    Dim emails As New Collection
    Dim email As New Dictionary
    
    email("email_address") = GetCellValue("Contacts", "email_address")
    email("primary_address") = True
    
    emails.add email
        
    Set GetEmails = emails
    
End Function

Public Function GetAccountName(ByVal defaultName As String) As String

    Dim accountName As New Dictionary
    
    If (defaultName = vbNullString) Then
        accountName("name") = GetCellValue("Accounts", "name")
    Else
        accountName("name") = defaultName
    End If
    
    GetAccountName = JsonConverter.ConvertToJson(accountName, Whitespace:=3)
    
End Function

Public Function GetOpportunity(ByVal opportunityId As String) As String

    Dim opportunity As New Dictionary
    
    Dim rowNum As Variant
    
    For Each rowNum In OpportunityRows
        Dim varName As String
        Dim varVal As String
        varName = CStr(Cells(CInt(rowNum), 6))
        varVal = CStr(Cells(CInt(rowNum), 3))
        
        If (varName <> "" And varName <> "id") Then
            opportunity(varName) = varVal
        End If
    Next rowNum
    
    If (opportunityId <> vbNullString And opportunityId <> "0") Then
        opportunity("id") = opportunityId
    End If
    
    GetOpportunity = JsonConverter.ConvertToJson(opportunity, Whitespace:=3)
    
End Function

Public Function GetQuote(ByVal quoteId As String, ByVal opportunityId As String, Optional ByVal products As Variant) As String

    Dim quote As New Dictionary
    Dim opp As New Dictionary
    Dim oppName As String
    oppName = GetCellValue("Opportunities", "name")
   
    quote("name") = oppName
    quote("opportunities_quotes_1_name") = oppName
    quote("opportunities_quotes_1opportunities_ida") = opportunityId
    
    opp("name") = oppName
    opp("id") = opportunityId
    
    Set quote("opportunities_quotes_1") = opp
    
    If (IsMissing(products) = False) Then
        Dim productBundles As New Dictionary
        Dim create As New Collection
        Dim productBundle1 As New Dictionary
        
        Dim productsQ As New Dictionary
        Dim add As New Collection
        Dim delete As New Collection
        Dim product1 As New Dictionary
        
        productBundle1("default_group") = True
        
        Dim p As Variant
        For Each p In products
            Dim prod As New Dictionary
            Set prod = New Dictionary
            prod("id") = p("id")
            If (p("quantity") = 0) Then
                delete.add prod
            Else
                add.add prod
            End If
        Next p
        
        If (add.count > 0) Then
            Set productsQ("add") = add
        End If
        
        If (delete.count > 0) Then
            Set productsQ("delete") = delete
        End If
        
        Set productBundle1("products") = productsQ
        
        create.add productBundle1
        
        Set productBundles("create") = create
        
        Set quote("product_bundles") = productBundles
    End If
    
    If (quoteId <> vbNullString And quoteId <> "0") Then
        quote("id") = quoteId
    End If
            
    GetQuote = JsonConverter.ConvertToJson(quote, Whitespace:=3)
    
End Function

Public Function GetQuotedLineItems(ByVal quoteId As String, ByVal accountId As String, ByVal opportunityId As String) As Variant

    Dim lineItems As New Collection
    
    Dim rowNum As Variant
    Dim lineItem As New Dictionary
    
    For Each rowNum In QuotedLineItemRows
    
        Dim partNumber As String
        partNumber = CStr(Cells(CInt(rowNum), 1))
        
        If (partNumber <> vbNullString And ApplicationProvider.StartsWith(partNumber, "Error") = False) Then
            Dim quantity As String
            quantity = CStr(Cells(CInt(rowNum), 6))
            
            Dim itemId As String
            
            itemId = GetLineItemId(partNumber)
            
            If (quantity <> vbNullString And quantity <> "n/a") Then
                If (CInt(quantity) > 0 Or itemId <> vbNullString) Then
                    Set lineItem = New Dictionary
                                
                    Dim partDescription As String
                    Dim listPrice As String
                    Dim price As String
                    
                    partDescription = CStr(Cells(CInt(rowNum), 2))
                    listPrice = CStr(Cells(CInt(rowNum), 3))
                    price = CStr(Cells(CInt(rowNum), 4))
                    
                    lineItem("name") = partNumber
                    lineItem("description") = partDescription
                    lineItem("list_price") = CDbl(listPrice)
                    lineItem("discount_price") = CDbl(price)
                    lineItem("discount_usdollar") = CDbl(price)
                    lineItem("cost_price") = CDbl(price)
                    lineItem("quantity") = CInt(quantity)
                    lineItem("quote_id") = quoteId
                    lineItem("account_id") = accountId
                    lineItem("opportunity_id") = opportunityId
                    
                    If (lineItem("quantity") = 0 And itemId <> vbNullString) Then
                        ExcelProvider.SetLineItemId partNumber, ""
                    End If
                    
                    If (itemId <> vbNullString) Then
                        lineItem("id") = itemId
                    End If
                                 
                    lineItems.add lineItem
                End If
            End If
        End If
    Next rowNum
    
    Set GetQuotedLineItems = lineItems
    
End Function

Public Function GetRowNumbers(ByVal sugarModule As String) As Variant

    Dim modRows As New Collection

    If (sugarModule = "Accounts") Then
        Set modRows = AccountRows
    ElseIf (sugarModule = "Contacts") Then
        Set modRows = ContactRows
    ElseIf (sugarModule = "Opportunities") Then
        Set modRows = OpportunityRows
    ElseIf (sugarModule = "Quotes") Then
        Set modRows = QuoteRows
    End If
    
    Set GetRowNumbers = modRows
    
End Function

Public Function GetRowNumber(ByVal sugarModule As String, ByVal sugarVarName As String) As Long

    Dim rowNum As Variant
    Dim rtnRowNum As Long
    
    For Each rowNum In GetRowNumbers(sugarModule)
        Dim varName As String
        varName = CStr(Cells(CInt(rowNum), 6))
        
        If (varName = sugarVarName) Then
            rtnRowNum = CLng(rowNum)
            Exit For
        End If
    Next rowNum
    
    GetRowNumber = rtnRowNum
    
End Function

Public Function GetLineItemRowNumber(ByVal sugarVarName As String) As Long

    Dim rowNum As Variant
    Dim rtnRowNum As Long
    
    For Each rowNum In QuotedLineItemRows
        Dim varName As String
        varName = CStr(Cells(CInt(rowNum), 1))
        
        If (varName = sugarVarName) Then
            rtnRowNum = CLng(rowNum)
            Exit For
        End If
    Next rowNum
    
    GetLineItemRowNumber = rtnRowNum
    
End Function

Public Function GetLineItemRowNumbers() As Variant

    Dim modRows As New Collection

    Dim i As Long
    For i = 1 To rows.count
        Dim modName As String
        modName = CStr(Cells(i, 8))
        If (modName = "Quoted Line Items") Then
            modRows.add i
            Dim qty As String
            qty = CStr(Cells(CInt(i), 6))
            If (qty <> vbNullString And qty <> "0") Then
                LineItemCount = LineItemCount + 1
            End If
        End If
    Next i
    
    Set GetLineItemRowNumbers = modRows
    
End Function

Public Function GetLineItemId(ByVal partNumber As String) As String

    Dim val As String
    val = Cells(GetLineItemRowNumber(partNumber), 9)
    GetLineItemId = CStr(val)
    
End Function

Public Function GetCellValue(ByVal sugarModule As String, ByVal sugarVarName As String) As String

    Dim rowNum As Variant
    Dim cellVal As String
    
    For Each rowNum In GetRowNumbers(sugarModule)
        Dim varName As String
        varName = CStr(Cells(CInt(rowNum), 6))
        
        If (varName = sugarVarName) Then
            cellVal = CStr(Cells(CInt(rowNum), 3))
            Exit For
        End If
    Next rowNum
    
    GetCellValue = cellVal
    
End Function

Public Function SetLineItemId(ByVal lineItemName As String, ByVal lineItemId As String) As Boolean

    If (lineItemName <> vbNullString) Then
        Cells(GetLineItemRowNumber(lineItemName), 9) = lineItemId
    End If
    
End Function

Public Function SetModuleId(ByVal moduleName As String, ByVal id As String)

    Cells(GetRowNumber(moduleName, "id"), 2) = id
    
End Function

Public Function SetVariables()
 
    Dim modRows As New Collection
    FathymFlowProvider.ExistingRecord = True

    Dim i As Long
    For i = 1 To rows.count
        Dim modName As String
        Dim modVar As String
        Dim modVal As String
        modName = CStr(Cells(i, 5))
        modVar = CStr(Cells(i, 6))
        modVal = CStr(Cells(i, 3))
        If (modName = "Accounts") Then
            AccountRows.add i
            If (modVar = "id") Then
                If (modVal <> vbNullString And modVal <> "0") Then
                    FathymFlowProvider.SugarAccountId = modVal
                Else
                    FathymFlowProvider.ExistingRecord = False
                End If
            End If
            If (modVar = "name") Then
                If (modVal = vbNullString Or modVal = "0") Then
                    ApplicationProvider.SetError ("Account Name cannot be blank")
                End If
            End If
        ElseIf (modName = "Contacts") Then
            ContactRows.add i
            If (modVar = "id") Then
                If (modVal <> vbNullString And modVal <> "0") Then
                    FathymFlowProvider.SugarContactId = modVal
                Else
                    FathymFlowProvider.ExistingRecord = False
                End If
            End If
            If (modVar = "name") Then
                If (modVal = vbNullString Or modVal = "0") Then
                    ApplicationProvider.SetError ("Contact Name cannot be blank")
                End If
            End If
        ElseIf (modName = "Opportunities") Then
            OpportunityRows.add i
            If (modVar = "id") Then
                If (modVal <> vbNullString And modVal <> "0") Then
                    FathymFlowProvider.SugarOpportunityId = modVal
                Else
                    FathymFlowProvider.ExistingRecord = False
                End If
            End If
            If (modVar = "name") Then
                If (modVal = vbNullString Or modVal = "0") Then
                    ApplicationProvider.SetError ("Project cannot be blank")
                End If
            End If
        ElseIf (modName = "Quotes") Then
            QuoteRows.add i
            If (modVar = "id") Then
                If (modVal <> vbNullString And modVal <> "0") Then
                    FathymFlowProvider.SugarQuoteId = modVal
                Else
                    FathymFlowProvider.ExistingRecord = False
                End If
            End If
        End If

    Next i
    
End Function
