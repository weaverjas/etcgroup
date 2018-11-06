Attribute VB_Name = "OnePageProjectSummary"


Sub PopulateDataFromRow()
    Dim N As Long, i As Long
    Dim strProjectInfo As String
    Dim strProjectName As String
    Dim strStart As String
    Dim strEnd As String
    Dim goodRow As Range
    
   'Make input sheet the active sheet
    Worksheets("input").Activate
    
    N = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To N
         'Make input sheet the active sheet
        Worksheets("input").Activate
        strProjectInfo = Cells(i, "A").Value
        If InStr(1, strProjectInfo, "Project Number") Then 'Check for relevent data
            'We have a useful row, lets make a new page
            strProjectName = GetProjectName(strProjectInfo)
            
            'Lets make a new page
            Module1.AddWorkSheetByName (Left(strProjectName, 31))
            
            'Time to poulate the new sheet
            CopyFromTo "C" & i, "B4", strProjectName 'Project Number
            CopyFromTo "Z" & i, "B6", strProjectName 'Project Manager
            CopyFromTo "O" & i, "B11", strProjectName 'Direct Labor Budget
            CopyFromTo "S" & i, "B12", strProjectName 'Direct Consultant Budget
            CopyFromTo "Q" & i, "B13", strProjectName 'Direct Expense Budget
            CopyFromTo "P" & i, "G11", strProjectName 'Direct Labor Cost
            CopyFromTo "T" & i, "G12", strProjectName 'Direct Consultant Cost
            CopyFromTo "R" & i, "G13", strProjectName 'Direct Expense Cost
            CopyFromTo "U" & i, "B14", strProjectName 'Margin Direct Labor pulling from Reimb
            CopyFromTo "G" & i, "A19", strProjectName 'Percentage of work complete
            CopyFromTo "AA" & i, "C19", strProjectName 'Hours used
            CopyFromTo "AB" & i, "J19", strProjectName 'AR Amount
            CopyFromTo "F" & i, "H19", strProjectName 'Project billed to date
        End If
    Next i
End Sub

'Return just the name of the project
Private Function GetProjectName(strInfo As String) As String
    Dim projectInfo() As String
    Dim garbo() As String
    Dim origStr As String
    Dim tx As Range

    garbo = Split(strInfo, ": ")
    projectInfo = Split(garbo(1), " ", 2)
    GetProjectName = Left(projectInfo(1), 31)
End Function

Private Function CopyFromTo(strFromCell As String, strToCell As String, strDestSheet As String)
    Dim strData As String
    
    Set Destsheet = Sheets(Left(strDestSheet, 31))
    Set SrcSheet = Sheets("input")
    strData = Sheets("input").Range(strFromCell).Value
    Sheets(strDestSheet).Range(strToCell).Value = strData
End Function
