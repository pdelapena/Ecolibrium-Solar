Sub OutlookToEcoCalcs()
    
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object
    Dim olItem As Outlook.MailItem
    Dim vText As Variant
    Dim sText As String
    Dim vItem As Variant
    Dim vItem1 As Variant
    Dim vAddr As Variant
    Dim oRng As Object
    Dim i As Long, j As Long
    Dim rCount As Long
    Dim sAddr As String
    Dim bXstarted As Boolean
    Dim cellAddress As String
    
    Dim uName As String
    Dim strPath As String
    Dim fileSaveName As String
    Dim templateName As String

    uName = Environ("UserName")
    
    strPath = "C:\Users\" & uName & "\Desktop\ECOCALCS_CURRENT"

    templateName = GetEcocalcsTemplate(strPath)

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "No Items selected!", vbCritical, "Error"
        Exit Sub
    End If
    
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    
    If Err <> 0 Then
        Application.StatusBar = "Please wait while Excel source is opened ... "
        Set xlApp = CreateObject("Excel.Application")
    End If
    On Error GoTo 0
    bXstarted = True
    'Open the workbook to input the data
     
    Set xlWB = xlApp.Workbooks.Open(strPath & "\" & templateName)
    Set xlSheet = xlWB.Sheets("IN-N-OUT")
    
    Dim id As Variant
    Dim accountIdentifiers As New Collection
    Set accountIdentifiers = GetSugarUniqueIdentifiers(xlSheet, "Accounts")
    
    For Each id In accountIdentifiers
        For Each olItem In Application.ActiveExplorer.Selection
            sText = olItem.Body
            vText = Split(sText, Chr(13))
            For i = UBound(vText) To 0 Step -1
                If InStr(1, vText(i), id) > 0 Then
                    vItem = Split(vText(i), Chr(58))
                    cellAddress = GetCellAddressByUniqueId(xlSheet, id)
                    xlSheet.Range(cellAddress) = Trim(vItem(1))
                End If
            Next i
        Next olItem
    Next id
    
    Dim contactIdentifiers As New Collection
    Set contactIdentifiers = GetSugarUniqueIdentifiers(xlSheet, "Contacts")
    
    For Each id In contactIdentifiers
        For Each olItem In Application.ActiveExplorer.Selection
            sText = olItem.Body
            vText = Split(sText, Chr(13))
            For i = UBound(vText) To 0 Step -1
                If InStr(1, vText(i), id) > 0 Then
                    vItem = Split(vText(i), Chr(58))
                    vItem1 = Replace(Trim(vItem(1)), " <tel", "")
                    vItem1 = Replace(Trim(vItem1), " <mailto", "")
                    cellAddress = GetCellAddressByUniqueId(xlSheet, id)
                    xlSheet.Range(cellAddress) = Trim(vItem1)
                End If
            Next i
        Next olItem
    Next id
    
    Dim oppIdentifiers As New Collection
    Set oppIdentifiers = GetSugarUniqueIdentifiers(xlSheet, "Opportunities")
    
    For Each id In oppIdentifiers
        For Each olItem In Application.ActiveExplorer.Selection
            sText = olItem.Body
            vText = Split(sText, Chr(13))
            For i = UBound(vText) To 0 Step -1
                If InStr(1, vText(i), id) > 0 Then
                    vItem = Split(vText(i), Chr(58))
                    vItem1 = Replace(Trim(vItem(1)), " <https", "")
                    cellAddress = GetCellAddressByUniqueId(xlSheet, id)
                    xlSheet.Range(cellAddress) = Trim(vItem1)
                End If
            Next i
        Next olItem
    Next id
    
    Dim quotedLineItemIdentifiers As New Collection
    Set quotedLineItemIdentifiers = GetPartNames(xlSheet, "Quoted Line Items")
    
    For Each id In quotedLineItemIdentifiers
        For Each olItem In Application.ActiveExplorer.Selection
            sText = olItem.Body
            vText = Split(sText, Chr(13))
            For i = UBound(vText) To 0 Step -1
                If InStr(1, vText(i), id) > 0 Then
                    vItem = Split(vText(i), Chr(58))
                    cellAddress = GetCellAddressByPartName(xlSheet, id)
                    xlSheet.Range(cellAddress) = Trim(vItem(1))
                End If
            Next i
        Next olItem
    Next id
    
    Dim questions As New Collection
    Set questions = GetSugarUniqueIdentifiers(xlSheet, "Questions")
    
    For Each id In questions
        For Each olItem In Application.ActiveExplorer.Selection
            sText = olItem.Body
            vText = Split(sText, Chr(13))
            For i = UBound(vText) To 0 Step -1
                If InStr(1, vText(i), id) > 0 Then
                    vItem = Split(vText(i), Chr(58))
                    cellAddress = GetCellAddressByUniqueId(xlSheet, id)
                    xlSheet.Range(cellAddress) = Trim(vItem(1))
                End If
            Next i
        Next olItem
    Next id

    fileSaveName = xlWB.Application.GetSaveAsFilename( _
            InitialFileName:="{customer}_{project}_" & templateName, _
            fileFilter:="Workbook (*.xlsm), *xlsm")
    
    'xlWB.SaveAs FileName:=fileSaveName
    If fileSaveName <> "False" Then
        '-- Save and Closethe file ----------------
        'Application.DisplayAlerts = False
        'ActiveWorkbook.SaveAs FileName:=sFileSaveName, _
                           FileFormat:=xlExcel8
        xlWB.SaveAs FileName:=fileSaveName
        'Application.ActiveWorkbook.Close
        'Application.DisplayAlerts = True
    Else
        '-- Close the file w/out saving it ---------
        xlWB.Close SaveChanges:=False
    End If
    
    If bXstarted Then
        xlApp.Quit
    End If
    
    Set xlApp = Nothing
    Set xlWB = Nothing
    Set xlSheet = Nothing
    Set olItem = Nothing
    
End Sub

Function FindAll(ByVal rng As Variant, ByVal searchTxt As String) As Collection
    Dim foundCell As Variant
    Dim firstAddress
    Dim rResult As New Collection
    
    With rng
        Set foundCell = .Find(What:=searchTxt)
        If Not foundCell Is Nothing Then
            firstAddress = foundCell.address
            Do
                rResult.Add CInt(foundCell.row)
                Set foundCell = .FindNext(foundCell)
            Loop While Not foundCell Is Nothing And foundCell.address <> firstAddress
        End If
    End With

    Set FindAll = rResult
End Function

Function FindAllFunctionValues(ByVal rng As Variant, ByVal searchTxt As String) As Collection
    Dim foundCell As Variant
    Dim firstAddress
    Dim rResult As New Collection
    
    With rng
        Set foundCell = .Find(What:=searchTxt, LookIn:=xlValues)
        If Not foundCell Is Nothing Then
            firstAddress = foundCell.address
            Do
                rResult.Add CInt(foundCell.row)
                Set foundCell = .FindNext(foundCell)
            Loop While Not foundCell Is Nothing And foundCell.address <> firstAddress
        End If
    End With

    Set FindAllFunctionValues = rResult
End Function

Function GetCellAddressByUniqueId(ByVal xlSheet As Object, ByVal xlRowId As String) As String
    Dim address As String
    Dim rows As Collection
    Set rows = FindAll(xlSheet.Range("A:A"), xlRowId)
    Dim rowNum As Variant
    Dim rng As String
    
    For Each rowNum In rows
        address = "B" & rowNum
        Exit For
    Next rowNum
    
    GetCellAddressByUniqueId = address
End Function

Function GetCellAddressByPartName(ByVal xlSheet As Object, ByVal xlRowId As String) As String
    Dim address As String
    Dim rows As Collection
    Set rows = FindAllFunctionValues(xlSheet.Range("B:B"), xlRowId)
    Dim rowNum As Variant
    Dim rng As String
    
    For Each rowNum In rows
        address = "E" & rowNum
        Exit For
    Next rowNum
    
    GetCellAddressByPartName = address
End Function

Function GetEcocalcsTemplate(inputDirectoryToScanForFile) As String
    Dim StrFile As String

    StrFile = Dir(inputDirectoryToScanForFile & "\*.xlsm")

    GetEcocalcsTemplate = StrFile
End Function

Function GetSugarUniqueIdentifiers(ByVal xlSheet As Object, ByVal sugarMod As String) As Collection
    Dim address As String
    Dim rows As Collection
    Dim ids As New Collection
    Set rows = FindAll(xlSheet.Range("E:E"), sugarMod)
    Dim rowNum As Variant
    Dim rng As String
    
    For Each rowNum In rows
        rng = "A" & rowNum
        ids.Add xlSheet.Range(rng).Value2
    Next rowNum
    
    Set GetSugarUniqueIdentifiers = ids
End Function

Function GetPartNames(ByVal xlSheet As Object, ByVal sugarMod As String) As Collection
    Dim address As String
    Dim rows As Collection
    Dim ids As New Collection
    Set rows = FindAll(xlSheet.Range("H:H"), sugarMod)
    Dim rowNum As Variant
    Dim rng As String
    
    For Each rowNum In rows
        rng = "B" & rowNum
        ids.Add xlSheet.Range(rng).Value2
    Next rowNum
    
    Set GetPartNames = ids
End Function
