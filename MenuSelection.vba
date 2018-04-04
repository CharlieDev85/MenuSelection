Private Sub Worksheet_Change(ByVal Target As Range)
    Dim SelectedState As String
    If Target.Address = "$H$8" Then
        SelectedState = Range("H8").Value
        Call SetListCounties(SelectedState)
    End If
End Sub

Sub SetListCounties(ByVal SelectedState As String)
    Dim TrimmedState As String
    Dim RangeName As String

    TrimmedState = Replace(SelectedState, " ", "")
    RangeName = "=counties" & TrimmedState
    Range("H9").Select
    Range("H9").Value = ""
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=RangeName
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Sub main()
    Application.ScreenUpdating = False
    Sheets("raw").Visible = True
    Dim State As String
    Dim County As String
    State = State1()
    County = County1()
    SelectState (State)
    SelectCounty (County)
    CopyPaste
    Sheets("raw").Visible = False
    Application.ScreenUpdating = True
End Sub

Public Function State1() As String
State1 = Range("H8").Value
End Function

Public Function County1() As String
County1 = Range("H9").Value
End Function

Sub SelectState(State As String)
    Sheets("raw").Visible = True
    Sheets("raw").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("State").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("State").CurrentPage = State
End Sub
Sub SelectCounty(County As String)
    Sheets("raw").Select
    ActiveSheet.PivotTables("PivotTable1").PivotFields("County").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("County").CurrentPage = County
End Sub

Public Function CopyPaste()
    Sheets("raw").Select
    ThisWorkbook.Sheets("raw").Range("Q37:AC50").Select
    Selection.Copy
    Sheets("MENU").Select
    Range("C20").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H8").Select
End Function

