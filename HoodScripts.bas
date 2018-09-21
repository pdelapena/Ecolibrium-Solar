Attribute VB_Name = "HoodScripts"
Function PopTheHood5D()
'
' PopTheHood5D Macro

    Columns("P:CN").Select
    Range("BG1").Activate
    Selection.EntireColumn.Hidden = False
    rows("8:202").Select
    Range("N8").Activate
    Selection.EntireRow.Hidden = False
    Range("A1").Select

End Function

Function CloseTheHood5D()
'
' CloseTheHood5D() Macro

    Union(Range( _
        "45:45,42:42,35:35,33:33,30:30,27:27,24:24,21:21,18:18,15:15,12:12,155:155,153:153,150:150,147:147,144:144,141:141,138:138,131:131,129:129,126:126,123:123,120:120,117:117,114:114,111:111,108:108,101:101,99:99,96:96,93:93,90:90" _
        ), Range("87:87,84:84,81:81,78:78,71:71,69:69,62:62,60:60,57:57,50:50,48:48")). _
        Select
    Selection.EntireRow.Hidden = True
    Columns("R:BF").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
End Function

Function CloseTheHoodPort2plus()
'
' CloseTheHoodPort2plus() Macro

    Union(Range( _
        "45:45,42:42,35:35,33:33,30:30,27:27,24:24,21:21,18:18,15:15,12:12,155:155,153:153,150:150,147:147,144:144,141:141,138:138,131:131,129:129,126:126,123:123,120:120,117:117,114:114,111:111,108:108,101:101,99:99,96:96,93:93,90:90" _
        ), Range("87:87,84:84,81:81,78:78,71:71,69:69,62:62,60:60,57:57,50:50,48:48")). _
        Select
    Selection.EntireRow.Hidden = True
    Columns("Y:CM").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
End Function

Function CloseTheHoodLand2plus()
'
' CloseTheHoodLand2plus()  Macro

    Union(Range( _
        "45:45,42:42,35:35,33:33,30:30,27:27,24:24,21:21,18:18,15:15,12:12,155:155,153:153,150:150,147:147,144:144,141:141,138:138,131:131,129:129,126:126,123:123,120:120,117:117,114:114,111:111,108:108,101:101,99:99,96:96,93:93,90:90" _
        ), Range("87:87,84:84,81:81,78:78,71:71,69:69,62:62,60:60,57:57,50:50,48:48")). _
        Select
    Selection.EntireRow.Hidden = True
    Columns("R:BM").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
End Function



