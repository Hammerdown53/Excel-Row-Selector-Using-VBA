VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub SelectRowsContainingText()
    Dim searchText As Variant
    Dim cell As Range
    Dim text As Variant
    Dim selectedRange As Range ' New line
    
    ' Define the array of texts to search for
    searchText = Array("", _
                       "", _
                       "", _
                       "", _
                       "", _
                       "", _
                       "", _
                       "", _
                       "", _
                       "", _
                       "", _
                       "", _
                       "", _
                       "", _
                       "")

    For Each cell In ActiveSheet.UsedRange
        ' Check if the cell contains any of the specified texts
        For Each text In searchText
            If InStr(cell.Value, text) > 0 Then
                ' If found, add the cell to the selected range
                If selectedRange Is Nothing Then
                    Set selectedRange = cell
                Else
                    Set selectedRange = Union(selectedRange, cell)
                End If
                Exit For
            End If
        Next text
    Next cell
    
    If Not selectedRange Is Nothing Then
        selectedRange.EntireRow.Select
    End If
End Sub

