Attribute VB_Name = "DataHelper"
' #TSQL Settings
Public Const SQL_COLUMN_QUAL_LEFT As String = "["
Public Const SQL_COLUMN_QUAL_RIGHT As String = "]"

' #Snowflake Settings
'Public Const SQL_COLUMN_QUAL_LEFT As String = """"
'Public Const SQL_COLUMN_QUAL_RIGHT As String = """"

Public Const JSON_QUOTE = "'" '""""

Public Const TAB_SPACE_COUNT As String = 2


'''''''''''''''''''
' Public Functions
'''''''''''''''''''


''
' Put text into clipboard ready to paste
' @param {String} txt
''
Public Function Helper_TxtToClibboard(txt As String)
    ' src: https://stackoverflow.com/questions/59706428/copy-text-to-clipboard-with-excel-365-vba
    Dim x: x = txt ' Cast to variant for 64-bit VBA support
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
                Case Len(txt)
                    .setData "text", x
                Case Else
                    Helper_TxtToClibboard = .GetData("text")
            End Select
        End With
    End With
End Function

''
' Turn range to filter clause(IN statement) for (RxC = Nx1) range or to WHERE clause for RxC = 2xN)
' e.g:
'   IN:    IN ('Code', 'FR01', 'FR02')
'   IN:    [Code] IN ('FR01', 'FR02')
'   WHERE: 1 = 1
'          AND [ProductID] = '1'
'          AND [Code] = 'FR01'
'          AND [Name] = 'Apple'
'
' @param {Range} rng
''
Public Function Helper_Range2SQL(rng As Range)
    If (rng Is Nothing) Then Exit Function
    If (rng.Rows.Count = 0 Or rng.Cells.Count = 0) Then Exit Function
    Dim TAB_SPACE: TAB_SPACE = Space(TAB_SPACE_COUNT)
    
    Dim IsAutoFilterOn: IsAutoFilterOn = IIf(Not ActiveSheet.AutoFilter Is Nothing, True, False)
    Dim IsFilterModeOn: IsFilterModeOn = ActiveSheet.FilterMode
    
    Dim filteredRng As Range
    Set filteredRng = IIf(IsAutoFilterOn And IsFilterModeOn, rng.SpecialCells(xlCellTypeVisible), rng)
    
    Dim retSQL As String, ii As Integer, col As String, val As String
    If filteredRng.Columns.Count > 1 And rng.Rows.Count > 1 Then
        retSQL = TAB_SPACE & "1 = 1 " & vbLf
        For ii = 1 To filteredRng.Columns.Count
            col = filteredRng.Cells(1, ii)
            val = rng.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1).Areas(1).Cells(1, ii) ' Cells in first row from filtered range
            If val = "NULL" Or val & "NULL" = "NULL" Then
                retSQL = retSQL & TAB_SPACE & "AND " & SQL_COLUMN_QUAL_LEFT & col & SQL_COLUMN_QUAL_RIGHT & " IS NULL " & vbLf
            Else
                retSQL = retSQL & TAB_SPACE & "AND " & SQL_COLUMN_QUAL_LEFT & col & SQL_COLUMN_QUAL_RIGHT & " = '" & val & "'" & vbLf
            End If
        Next ii
    ElseIf rng.Columns.Count = 1 And rng.Rows.Count > 1 Then
        ' src: https://vba2008.wordpress.com/2008/10/17/check-if-range-is-filtered-check-if-sheet-has-autofilter-using-excel-vba/
        
        
        If Not IsAutoFilterOn Then ' NOT on filter mode of this column
            retSQL = TAB_SPACE & "IN ("
            For Each xell In filteredRng
                val = xell.Value
                If val = "NULL" Or val & "NULL" = "NULL" Then
                    ' skip this row
                Else
                    retSQL = retSQL & "'" & val & "', "
                End If
            Next xell
        ElseIf IsAutoFilterOn Then ' filter is on this filter or rows are filtered by this column
            retSQL = SQL_COLUMN_QUAL_LEFT & rng.Cells(1, 1) & SQL_COLUMN_QUAL_RIGHT & " IN ("
            ii = 0
            
            For Each xell In filteredRng
                If ii > 0 Then
                    val = xell.Value
                    If val = "NULL" Or val & "NULL" = "NULL" Then
                        ' skip this row
                    Else
                        retSQL = retSQL & "'" & val & "', "
                    End If
                End If
                ii = ii + 1
            Next xell
            
        End If
        retSQL = Left(retSQL, Len(retSQL) - 2) & ")"
    Else
    End If
    
    Helper_Range2SQL = retSQL
End Function


''
' Turn a range to JSON literal
' @param {Range} rng
''
Public Function Helper_Range2JSON(rng As Range) As String
    If (rng Is Nothing) Then Exit Function
    If (rng.Rows.Count = 0 Or rng.Cells.Count = 0) Then Exit Function
    
    ' src: http://niraula.com/blog/convert-excel-data-json-format-using-vba/
    ' Make sure there are two columns in the range
    If rng.Columns.Count < 2 Then
        ToJSON = CVErr(xlErrNA)
        Exit Function
    End If

    Dim dataLoop, headerLoop As Long
    ' Get the first row of the range as a header range
    Dim headerRange As Range: Set headerRange = Range(rng.Rows(1).Address)
    
    ' We need to know how many columns are there
    Dim colCount As Long: colCount = headerRange.Columns.Count
    
    Dim json As String: json = "["
    
    For dataLoop = 1 To rng.Rows.Count
        ' Skip the first row as it's been used as a header
        If dataLoop > 1 Then
            ' Start data row
            Dim rowJson As String: rowJson = "{"
            
            ' Loop through each column and combine with the header
            For headerLoop = 1 To colCount
                rowJson = rowJson & JSON_QUOTE & headerRange.Value2(1, headerLoop) & JSON_QUOTE & ":"
                rowJson = rowJson & JSON_QUOTE & rng.Value2(dataLoop, headerLoop) & JSON_QUOTE
                rowJson = rowJson & ","
            Next headerLoop
            
            ' Strip out the last comma
            rowJson = Left(rowJson, Len(rowJson) - 1)
            
            ' End data row
            json = json & rowJson & "},"
        End If
    Next
    
    ' Strip out the last comma
    json = Left(json, Len(json) - 1)
    
    json = json & "]"
    
    Helper_Range2JSON = json
End Function
