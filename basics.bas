Attribute VB_Name = "DoSomethingWithCells"
Sub CalculateGreen()
    ' Making variables and giving them types
    Dim selectedRange As Range
    Dim cell As Range
    Dim greenCellCount As Integer
    
    ' Make chosen cells to cells used
    Set selectedRange = Selection
    
    ' Test if any cells have been chosen
    If selectedRange Is Nothing Then
        MsgBox "Valitse solut"
        Exit Sub
    End If
    
    ' Giving variable a value

    greenCellCount = 0
    
    ' Loop through every cell in selected range and if RGB matches make the count + 1
    For Each cell In selectedRange
        If cell.Interior.Color = RGB(146, 208, 80) Then
            greenCellCount = greenCellCount + 1
        End If

    ' Go to next cell
    Next cell

    MsgBox "Greens: " & greenCellCount
End Sub

Sub nextCell()

    ' Making variables and giving them types
    Dim selectedRange As Range
    Dim cell As Range
    Dim nextCell As Range
    
    ' Make chosen cells to cells used
    Set selectedRange = Selection
    
    ' Test if any cells have been chosen
    If selectedRange Is Nothing Then
        MsgBox "Valitse solut"
        Exit Sub
    End If
    
    ' Loops through every cell in selected range and if RGB matches, make the next columns cells value to 1
    For Each cell In selectedRange
        If cell.Interior.Color = RGB(146, 208, 80) Then
            Set nextCell = cell.Offset(0, 1)
            nextCell.Value = 1
        End If

    ' Go to next cell
    Next cell
        
End Sub