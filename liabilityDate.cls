Sub HideColumns()
' HideColumns Macro
    
    Application.ScreenUpdating = False
    
    Columns("D:D").Select
    Selection.EntireColumn.Hidden = True
    Columns("G:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("H:H").Select
    Selection.EntireColumn.Hidden = True
    Columns("I:I").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:J").Select
    Selection.EntireColumn.Hidden = True
    Columns("L:L").Select
    Selection.EntireColumn.Hidden = True
    Columns("M:M").Select
    Selection.EntireColumn.Hidden = True
    
    ActiveWindow.SmallScroll Down:=-10000
    ActiveSheet.Range("Z1").Select
    Application.ScreenUpdating = True
End Sub
Sub ShowColumns()
' ShowColumns Macro
    Application.ScreenUpdating = False
    
    Columns("D:D").Select
    Selection.EntireColumn.Hidden = False
    Columns("G:G").Select
    Selection.EntireColumn.Hidden = False
    Columns("H:H").Select
    Selection.EntireColumn.Hidden = False
    Columns("I:I").Select
    Selection.EntireColumn.Hidden = False
    Columns("J:J").Select
    Selection.EntireColumn.Hidden = False
    Columns("L:L").Select
    Selection.EntireColumn.Hidden = False
    Columns("M:M").Select
    Selection.EntireColumn.Hidden = False
    
    ActiveWindow.SmallScroll Down:=-10000
    ActiveSheet.Range("Z1").Select
    Application.ScreenUpdating = True
End Sub
Sub liabilityDate()
' Programmer: Abraham Juliot
' Macro: liabilityDate
    
    ' Get a valid date from the user
    Dim actualDate As Date 'hold the actual date
    actualDate = Date$ 'assign today's date to actualDate
    userDate = InputBox( _
        "Enter a valid date." _
        & vbNewLine _
        & vbNewLine & "For Example:" _
        & vbNewLine _
        & vbNewLine & Format(actualDate, "mmmm d") _
        & vbNewLine & Format(actualDate, "mmm d") _
        & vbNewLine & Format(actualDate, "d mmm") _
        & vbNewLine & Format(actualDate, "mm/dd/yy") _
        & vbNewLine & Format(actualDate, "mm/dd/yyyy") _
        & vbNewLine & Format(actualDate, "dd-mm") _
        & vbNewLine & Format(actualDate, "mm-dd") _
        & vbNewLine _
        & vbNewLine & "Make sure you enter the year to view a previous or future year." _
    )
    
    ' Validate the user's input
    'if nothing was entered, then end the program
    If userDate = "" Then
    End
    'else if userDate is not a date, then alert the user and end the program
    ElseIf Not IsDate(userDate) Then
    MsgBox "Invalid date": End
    End If
    
    ' Clear filters before query begins
    Run "NewClearFilter"
    
    ' Turn off screen updating after InputBox is called and input is valid
    Application.ScreenUpdating = False
    
    ' Assign the date to query to the viewDate cell
    Range("viewDate").Value = Format(userDate, "mm/dd/yyyy")
    
    ' Assign the date to query to dateLastIssued and dateLastReturned
    Dim dateLastIssued As Date, dateLastReturned As Date
    dateLastIssued = userDate
    dateLastReturned = userDate
    
    ' Only show issue dates on and before the query date
    ' These deposits are received as of the query date
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5, Criteria1:= _
        "<=" & dateLastIssued
        
    ' Only show returned dates after the query date
    ' These deposits are not yet refunded or forfeited as of the query date
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=14, Criteria1:= _
        ">" & dateLastReturned, Operator:=xlOr, Criteria2:="="
        
    ' Clear any previous sorts and sort by date issued (ascending)
    Range("Table1[[#Headers],[Date Issued]]").Select
    
    ActiveWorkbook.Worksheets("Key Log").ListObjects("Table1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Key Log").ListObjects("Table1").Sort.SortFields.Add _
        Key:=Range("Table1[[#Headers],[Date Issued]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
        
    ActiveWorkbook.Worksheets("Key Log").ListObjects("Table1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Key Log").ListObjects("Table1").Sort.SortFields.Add _
        Key:=Range("Table1[[#Headers],[#Data],[Date Issued]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Key Log").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Only show rows with a positive deposit value ($5.00 or $10.00)
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=6, Criteria1:= _
    "=$5.00", Operator:=xlOr, Criteria2:="=$10.00"
    
    ' Scroll near the end of the report
    ActiveSheet.Cells(ActiveSheet.Rows.Count, Selection.Column).End(xlUp).Select
    ActiveWindow.ScrollRow = Selection.Row
    ActiveWindow.SmallScroll Down:=-50
    
    ' Turn on screen updating before MsgBox is called
    Application.ScreenUpdating = True
    
    ' Display the liability amount (as of the query date) to the user
    Dim liabilityAmount As Range
    Dim liabilityDate As Date
    Set liabilityAmount = Range("totalOwed")
    liabilityDate = userDate
    
    
    ' Display the liability amount (as of the query date) to the user
    Dim report As Integer
    
    report = MsgBox("The Liability amount as of " _
        & Format(liabilityDate, "mmmm dd, yyyy") _
        & " is " & Format(liabilityAmount, "$#,##0.00") & "." _
        & vbNewLine _
        & vbNewLine & "This report totals the key deposits " _
        & "received (and not yet refunded or forfeited) as of " _
        & Format(liabilityDate, "mmmm dd, yyyy.") _
        & vbNewLine _
        & vbNewLine & "Do you wish to analyze this report?", vbYesNo, _
        Format(liabilityDate, "mmmm dd, yyyy") & " Report" _
    )
    
    ' If no, then exit the query by clearing the filters
    If report = vbNo Then
    Run "NewClearFilter"
    ' Else, scroll to the beginning of the report
    Else
    Application.ScreenUpdating = False
    ActiveWindow.SmallScroll Down:=-10000
    
    ' Add report summary to text box
    ActiveSheet.Shapes.Range(Array("TextBox 5")).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = _
        "The Liability amount as of " _
        & Format(liabilityDate, "mmmm dd, yyyy") _
        & " is " & Format(liabilityAmount, "$#,##0.00") & "." _
        & vbNewLine _
        & vbNewLine & "You are viewing the total key deposits " _
        & "received (and not yet refunded or forfeited) as of " _
        & Format(liabilityDate, "mmmm dd, yyyy.")
    
    ' Outline the text box
    Application.CommandBars("Format Object").Visible = False
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(51, 102, 255)
        .Weight = 0.25
    End With
    Application.ScreenUpdating = True
    End If
    
    Application.ScreenUpdating = False
    Range("E3").Select
    Application.ScreenUpdating = True
               
End Sub
