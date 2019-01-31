
-------------------------------Identifying Active twitter Members for the day -------------------------------------------------------------------------------------------
Sub Active_List()

Dim i As Long, n As Long, ID As Long
Dim str1 As String, str2 As String
Sheets("ACTIVE MEMBERS").Range("A2:E104876").Clear
Application.ScreenUpdating = False
n = 2
For i = 1 To Sheets("SBU wise analysis").Range("A1048576").End(xlUp).Row
    
    If Sheets("SBU wise analysis").Range("F" & i) = "Sterlite" Then
        Sheets("ACTIVE MEMBERS").Range("A" & n) = Sheets("SBU wise analysis").Range("B" & i)
        Sheets("ACTIVE MEMBERS").Range("B" & n) = Sheets("SBU wise analysis").Range("C" & i)
        Sheets("ACTIVE MEMBERS").Range("D" & n) = Sheets("SBU wise analysis").Range("D" & i)
        Sheets("ACTIVE MEMBERS").Range("C" & n) = _
        Application.WorksheetFunction.VLookup(Sheets("ACTIVE MEMBERS").Range("B" & n), Sheets("master employee list").Range("C:D"), 2, False)
        Sheets("ACTIVE MEMBERS").Range("E" & n) = Application.WorksheetFunction.CountIf(Sheets("SBU wise analysis").Range("B:B"), Sheets("ACTIVE MEMBERS").Range("A" & n))
        n = n + 1
    End If
Next
Sheets("ACTIVE MEMBERS").Activate
Sheets("ACTIVE MEMBERS").Columns("A:E").EntireColumn.Select
Sheets("ACTIVE MEMBERS").Range("$A$1:$E$104876").RemoveDuplicates Columns:=2, Header:=xlYes

End Sub()
--------------------------------------Inactive Members--------------------------------------------------------------------------------------------------
Sub Inactive()

Dim i As Integer, n As Integer, ID As Integer

Sheets("INACTIVE MEMBERS").Range("A2:E104876").Clear
Application.ScreenUpdating = False
n = 2
For i = 2 To Sheets("master employee list").Range("A1048576").End(xlUp).Row
    On Error GoTo DoThis
    ID = Application.WorksheetFunction.VLookup(Sheets("master employee list").Range("C" & i), Sheets("ACTIVE MEMBERS").Range("B:D"), 1, False)
    Next
Exit Sub
DoThis:
    Sheets("INACTIVE MEMBERS").Range("A" & n) = Sheets("master employee list").Range("B" & i)
    Sheets("INACTIVE MEMBERS").Range("B" & n) = Sheets("master employee list").Range("C" & i)
    Sheets("INACTIVE MEMBERS").Range("C" & n) = Sheets("master employee list").Range("D" & i)
    Sheets("INACTIVE MEMBERS").Range("D" & n) = Sheets("master employee list").Range("G" & i)
    Sheets("INACTIVE MEMBERS").Range("E" & n) = Sheets("master employee list").Range("H" & i)
    n = n + 1
    
Resume Next
Application.ScreenUpdating = True
----------------------------------------------------Macro to remove duplicates-----------------------------------------------------------
Sub REMOVING_DUPLICATES()

REMOVING_DUPLICATES Macro



    Sheets("ACTIVE MEMBERS").Columns("A:E").EntireColumn.Select
    Sheets("ACTIVE MEMBERS").Range("$A$1:$E$104876").RemoveDuplicates Columns:=2, Header:= _
        xlYes
End Sub

-------------------------Inactive HoDs--------------------------------------------------------------------------------------------------
Sub Inactive_HoDs()
Dim i As Integer, n As Integer, ID As Double, Status As String
Sheets("INACTIVE_HoDs").Range("A2:E104876").Clear
Application.ScreenUpdating = False
n = 2
For i = 2 To Sheets("INACTIVE MEMBERS").Range("A1048576").End(xlUp).Row
    On Error Resume Next
    ID = Application.WorksheetFunction.VLookup(Sheets("INACTIVE MEMBERS").Range("B" & i), Sheets("reporting heads").Range("A:E"), 1, False)
If (Sheets("INACTIVE MEMBERS").Range("B" & i) = ID) Then
    Sheets("INACTIVE_HoDs").Range("A" & n) = Sheets("INACTIVE MEMBERS").Range("B" & i)
    Sheets("INACTIVE_HoDs").Range("B" & n) = Sheets("INACTIVE MEMBERS").Range("C" & i)
    Sheets("INACTIVE_HoDs").Range("C" & n) = Sheets("INACTIVE MEMBERS").Range("D" & i)
    Sheets("INACTIVE_HoDs").Range("D" & n) = Sheets("INACTIVE MEMBERS").Range("E" & i)
    Status = Application.WorksheetFunction.VLookup(Sheets("INACTIVE_HoDs").Range("A" & n), Sheets("master employee list").Range("C:I"), 7, False)
    Sheets("INACTIVE_HoDs").Range("E" & n) = Status
    n = n + 1
End If
Next
Application.ScreenUpdating = True
End Sub





