Attribute VB_Name = "Module1"
Sub EDIT_Unhide_All_Tabs()
Dim ws As Worksheet
Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Worksheets
        
        ws.Visible = xlSheetVisible
    
    Next ws
Application.ScreenUpdating = True
Worksheets("Change Log").Activate

End Sub
Sub EDIT_Unhide_Matrix_Tabs()
Application.ScreenUpdating = False
    For Each wsSheet In Worksheets
    ' If Statement
        If wsSheet.Name = "Land Coeffs 10deg" Or _
            wsSheet.Name = "Land Calcs 10deg (lift)" Or _
            wsSheet.Name = "Land Calcs 10deg (sliding)" Or _
            wsSheet.Name = "Land Coeffs 5deg" Or _
            wsSheet.Name = "Land Calcs 5deg (lift)" Or _
            wsSheet.Name = "Land Calcs 5deg (sliding)" Or _
            wsSheet.Name = "Port Coeffs" Or _
            wsSheet.Name = "Port Calcs (lift)" Or _
            wsSheet.Name = "Port Calcs (sliding)" Then
        wsSheet.Visible = True
    End If
    
    Next wsSheet

Application.ScreenUpdating = True
Worksheets("Land Matrix 10deg (lift)").Activate

End Sub
Sub EDIT_Hide_Background_Tabs()
Application.ScreenUpdating = False
    For Each wsSheet In Worksheets
    ' If Statement
        If wsSheet.Name = "Change Log" Or _
            wsSheet.Name = "IN-N-OUT" Or _
            wsSheet.Name = "EcoMount Inputs" Or _
            wsSheet.Name = "factors" Or _
            wsSheet.Name = "Land Coeffs 10deg" Or _
            wsSheet.Name = "Land Calcs 10deg (lift)" Or _
            wsSheet.Name = "Land Calcs 10deg (sliding)" Or _
            wsSheet.Name = "Land Coeffs 5deg" Or _
            wsSheet.Name = "Land Calcs 5deg (lift)" Or _
            wsSheet.Name = "Land Calcs 5deg (sliding)" Or _
            wsSheet.Name = "Port Coeffs" Or _
            wsSheet.Name = "Port Calcs (lift)" Or _
            wsSheet.Name = "Port Calcs (sliding)" Or _
            wsSheet.Name = "Area Reduction" Or _
            wsSheet.Name = "Friction Data" Or _
            wsSheet.Name = "SEAOC PV2 Calcs" Or _
            wsSheet.Name = "Seismic Attached Per-Array" Or _
            wsSheet.Name = "Seismic Calcs (Attached)" Or _
            wsSheet.Name = "Seismic Calcs (Unattached)" Or _
            wsSheet.Name = "Seismic Data (Unattached)" Or _
            wsSheet.Name = "snow load" Or _
            wsSheet.Name = "Strut Values" Or _
            wsSheet.Name = "Thermal Sliding Calcs" Or _
            wsSheet.Name = "Uplift Attachment Calcs" Or _
            wsSheet.Name = "ZIPS" Then
        wsSheet.Visible = False
    Else: wsSheet.Visible = True
    End If
    
    Next wsSheet

Application.ScreenUpdating = True
Worksheets("1-Eng Inputs").Activate

End Sub

Sub EDIT_Unhide_Matrix_Backgrounds()
Application.ScreenUpdating = False
    For Each wsSheet In Worksheets
    ' If Statement
        If wsSheet.Name = "Land Matrix 10deg (lift)" Or _
            wsSheet.Name = "Land Matrix 5deg (lift)" Or _
            wsSheet.Name = "Port Matrix (lift)" Then
        Worksheets(wsSheet.Name).Activate
        Call PopTheHood5D
            
        End If
    
    Next wsSheet

Application.ScreenUpdating = True
Worksheets("Land Matrix 10deg (lift)").Activate

End Sub
Sub EDIT_Hide_Matrix_Backgrounds()
Application.ScreenUpdating = False
    For Each wsSheet In Worksheets
    ' If Statement
        Select Case wsSheet.Name
            Case Is = "Land Matrix 10deg (lift)"
                Worksheets(wsSheet.Name).Activate
                Call CloseTheHoodLand2plus
                
            Case Is = "Land Matrix 5deg (lift)"
                Worksheets(wsSheet.Name).Activate
                Call CloseTheHood5D
                
            Case Is = "Port Matrix (lift)"
                Worksheets(wsSheet.Name).Activate
                Call CloseTheHoodPort2plus
        End Select
    
    Next wsSheet

Application.ScreenUpdating = True
Worksheets("Land Matrix 10deg (lift)").Activate
End Sub

