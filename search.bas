Attribute VB_Name = "search"
Option Explicit

Public Function getLastRow(worksheet As String) As Long
' PURPOSE: Finds the last row of the <worksheet>
' PARAMETERS: <worksheet> The name of the worksheet
' RETURNS: the row index of the last row
    
    Dim sheet As worksheet
    Set sheet = ThisWorkbook.Worksheets(worksheet)
    getLastRow = sheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

End Function

Public Function getLastCol(worksheet As String) As Long
' PURPOSE: Finds the last column of the <worksheet>
' PARAMETERS: <worksheet> The name of the worksheet
' RETURNS: the column index of the last column

    Dim sheet As worksheet
    Set sheet = ThisWorkbook.Worksheets(worksheet)
    getLastCol = sheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column

End Function

Public Function searchItem(item As String, worksheet As String, startPos As Long, searchType As Boolean, _
                           matchCase As Boolean, returns As Boolean) As Long
' PURPOSE: Finds the index of the <item> you are looking for
' PARAMETERS: <item> The word that is being looked for
'             <worksheet> The name of the worksheet
'             <startPos> Starting position of search
'             <searchType> True looks for cell with only full <item>, false looks for the cell that has <item>
'             <matchCase> True case sensitive, false otherwise
'             <returns> True for row, false for column
' RETURNS: the index of the <item> we are looking for, returns -1 if not found

    Dim sheet As worksheet
    Set sheet = ThisWorkbook.Worksheets(worksheet)
    Dim searchDecision As Integer
    
    If searchType Then
        searchDecision = xlWhole
    Else
        searchDecision = xlPart
    End If
    
    On Error GoTo notFound
        If returns Then
            searchItem = sheet.Cells.Find(What:=item, After:=Cells(startPos, 1), LookIn:=xlFormulas2, _
                         LookAt:=searchDecision, SearchOrder:=xlByRows, _
                         SearchDirection:=xlNext, matchCase:=matchCase, SearchFormat:=False).Row
        Else
            searchItem = sheet.Cells.Find(What:=item, After:=Cells(1, startPos), LookIn:=xlFormulas2, _
                         LookAt:=searchDecision, SearchOrder:=xlByColumns, _
                         SearchDirection:=xlNext, matchCase:=matchCase, SearchFormat:=False).column
        End If

Exit Function
notFound:
    searchItem = -1

End Function

Public Function main()
' PURPOSE: Tests the functions created!
    Debug.Print ("Row:" & getLastRow("Sheet1"))
    Debug.Print ("Col:" & getLastCol("Sheet1"))
    Debug.Print ("Search:" & searchItem("hello", "Sheet1", 1, True, True, True))
End Function
