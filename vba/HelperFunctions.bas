Attribute VB_Name = "HelperFunctions"

Public Function AddWorkSheetByName(strName As String)
    ' Test if worksheet already exists, create if not.
    CreateNewSheet strName
End Function

'Create a new worksheet only if a worksheet of that name does not exist
Private Function CreateNewSheet(strSheetName As String)
    If Not SheetExists(strSheetName) Then
        'Create a new worksheet and add it to the end
        Dim wrkSheet As Worksheet
        Set wrkSheet = Sheets.Add(After:=Sheets(Sheets.Count))
        'Set new worksheet name
        wrkSheet.name = Left(strSheetName, 31) 'Max Name Lenght
        CreateSheetLayout Left(strSheetName, 31)
    End If
End Function

'Setup a new worksheet with the predefined crap you wanted
Private Function CreateSheetLayout(strSheetName As String)
    Dim actSheet As Worksheet
    Dim srcRange As Range
    Dim fillRange As Range
    Dim N As String
    
    'Grab the new sheet
    Set actSheet = Sheets(Left(strSheetName, 31))
    
    'Stretch out some columns
    Columns("A:A").ColumnWidth = 16.67
    Columns("B:B").ColumnWidth = 15.83
    Columns("C:C").ColumnWidth = 15.83
    Columns("E:E").ColumnWidth = 10.33
    Columns("F:F").ColumnWidth = 10.5
    Columns("G:G").ColumnWidth = 6.67
    Columns("H:H").ColumnWidth = 8.33
    Columns("I:I").ColumnWidth = 8.33
    Columns("J:J").ColumnWidth = 10.5
    Columns("K:K").ColumnWidth = 8.33
    Columns("L:L").ColumnWidth = 19.17
    
    'And Stretch out some rows
    Rows("10:10").RowHeight = 32
    Rows("10:10").RowHeight = 32
    
    'Setup a GreyBar Titled One Page Project Summary Report starting at A3 to F3
    Set srcRange = actSheet.Range("A3")
    Set fillRange = actSheet.Range("A3:F3")
    CreateGreyBarWithTitle srcRange, fillRange, "One Page Project Summary Report"
    
    'Setup a Grey Bar titled Project from A9 to J9
    Set srcRange = actSheet.Range("A9")
    Set fillRange = actSheet.Range("A9:J9")
    CreateGreyBarWithTitle srcRange, fillRange, "Project"
    
    'Add a double grey bar
    CreateDoubleBorderByRange "A16:J16"
    
    'Add a solid single line
    CreateSingleBorderByRange "B10:L10"
    
    'start adding titles
    CreateBoldTextInCell "B10", "Original Budget    "
    CreateBoldTextInCell "D10", "EAC"
    CreateBoldTextInCell "E10", "PM EAC"
    CreateBoldTextInCell "G10", "JTD Cost"
    CreateBoldTextInCell "H10", "Remaining Budget"
    CreateBoldTextInCell "I10", "Margin"
    CreateBoldTextInCell "J10", "Margin Percent"
    
    'And still more titles
    CreateBoldTextInCell "A18", "Percentage Scope Complete"
    CreateBoldTextInCell "B18", "Hours remaining"
    CreateBoldTextInCell "C18", "Hours Used"
    CreateBoldTextInCell "H18", "Billings to date"
    CreateBoldTextInCell "J18", "Aged AR"
    
    'Almost done :)
    CreateTextInCell "A4", "Project Number"
    CreateTextInCell "A5", "Project Name"
    CreateTextInCell "A6", "Project Manager"
    CreateTextInCell "A11", "Direct Labor"
    CreateTextInCell "A12", "Direct Consultants"
    CreateTextInCell "A13", "Direct Expenses"
    CreateTextInCell "A14", "Reimbursable"
    CreateTextInCell "A15", "Total"
    
    'AddingFormualas
     CreateTextInCell "H11", "=B11-G11"
     
End Function

'Create a grey bar of a specified size with a given title displayd in black
Private Function CreateGreyBarWithTitle(srcRange As Range, fillRange As Range, strTitle As String)
    srcRange.Select
    Selection.AutoFill Destination:=fillRange, Type:=xlFillDefault
    fillRange.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Font
        .FontStyle = "Bold"
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.0999786370433668 'Changing this number will change the color of the bar
        .PatternTintAndShade = 0
    End With
    ActiveCell.FormulaR1C1 = strTitle
End Function

'Place non bold text in a specified cell
Private Function CreateTextInCell(strCell As String, strName As String)
    Dim targetCell As Range
    
    Set targetCell = Range(strCell)
    targetCell.Select
     ActiveCell.FormulaR1C1 = strName
    
End Function

'Place Bold Text in a specified cell
Private Function CreateBoldTextInCell(strCell As String, strName As String)
    Dim targetCell As Range
    
    Set targetCell = Range(strCell)
    targetCell.Select
    With Selection
        .WrapText = True
        .ShrinkToFit = False
        .MergeCells = False
    End With
    With Selection.Font
        .FontStyle = "Bold"
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ActiveCell.FormulaR1C1 = strName
End Function

Private Function CreateSingleBorderByRange(fillRange As String)
    Range(fillRange).Select
    Selection.Merge
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlSingle
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With

End Function

Private Function CreateDoubleBorderByRange(fillRange As String)
    Range(fillRange).Select
    Selection.Merge
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
End Function

'Test if a sheet exists
'Returns True if exists and False if not
Private Function SheetExists(strSheetName As String) As Boolean
    Dim obj As Object
    On Error GoTo HandleError
    Set obj = Sheets(strSheetName)
    SheetExists = True
    Exit Function
HandleError: ' obj will throw an error if the sheet name is invalid
    SheetExists = False
End Function
