Attribute VB_Name = "Module1"
Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function to check if a value is in an array of values
'INPUT: Pass the function a value to search for and an array of values of any data type.
'OUTPUT: True if is in array, false otherwise
Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function

Sub ChangeColorOnCondition(target As Range, variable As Single, Optional precision As Single = 0.0001)
     
        If Abs(target.Value - variable) > precision Then
                target.Interior.ColorIndex = 3
        Else
                target.Interior.Color = 16777215
        End If
   
End Sub

Sub Vectors()
      ' #################### IMPORTANT #######################
    ' A good idea for ALL labs (because at some point in time things could change) is creating variables to refer to the position of
    ' tables (it could be the left upper cell for instance) and reference all the calculations in that table to those coordinates.
    ' That way when a change is made in a sheet, the only task one has to do is to adjust that pair of variables to the new ones and
    ' the code will remain functional.
    Dim aX As Single
    Dim aY As Single
    Dim bX As Single
    Dim bY As Single
    Dim i As Single
    Dim j As Single
    Dim aDOTb As Single
    Dim theta As Single
    Dim dotProd As Single
    Dim Normal As Boolean
    Dim posAnswer(2) As String
    Dim negAnswer(2) As String
    Dim ws As Worksheet
    Dim ws1 As Worksheet
    Set ws = Worksheets("Tables")
    Set ws1 = Worksheets("Questions")
    
    'Defining variables related to the upper left corner of a table to make all the parameters relative to that point.
    Dim t2FirstRow As Single
    Dim t2FirstCol As Single
    
    t2FirstRow = 42
    t2FirstCol = 1
    
    aX = 7 * Cos(30 * WorksheetFunction.pi / 180)
    aY = 7 * Sin(30 * WorksheetFunction.pi / 180)
    bX = -4 * Cos(45 * WorksheetFunction.pi / 180)
    bY = -4 * Sin(45 * WorksheetFunction.pi / 180)
    
    aDOTb = aX * bX + aY * bY
    theta = WorksheetFunction.Degrees(WorksheetFunction.Acos(aDOTb / (AbsVector(aX, aY) * AbsVector(bX, bY))))
    
    posAnswer(1) = "Y"
    posAnswer(2) = "YES"
    negAnswer(1) = "N"
    negAnswer(2) = "NO"
    
    'Checking out x and y coord of the vectors
    ChangeColorOnCondition ws.Cells(11, 2), aX, 0.1
    ChangeColorOnCondition ws.Cells(12, 2), aY, 0.1
    ChangeColorOnCondition ws.Cells(11, 3), bX, 0.1
    ChangeColorOnCondition ws.Cells(12, 3), bY, 0.1
    ChangeColorOnCondition ws.Cells(11, 4), aX + bX, 0.1
    ChangeColorOnCondition ws.Cells(12, 4), aY + bY, 0.1
    ChangeColorOnCondition ws.Cells(11, 5), aX - bX, 0.1
    ChangeColorOnCondition ws.Cells(12, 5), aY - bY, 0.1
    ChangeColorOnCondition ws.Cells(11, 6), bX - aX, 0.1
    ChangeColorOnCondition ws.Cells(12, 6), bY - aY, 0.1
    ChangeColorOnCondition ws.Cells(11, 7), aX + 2 * bX, 0.1
    ChangeColorOnCondition ws.Cells(12, 7), aY + 2 * bY, 0.1
    ChangeColorOnCondition ws.Cells(11, 8), aX - 2 * bX, 0.1
    ChangeColorOnCondition ws.Cells(12, 8), aY - 2 * bY, 0.1
    ChangeColorOnCondition ws.Cells(11, 9), 2 * aX + 0.5 * bX, 0.1
    ChangeColorOnCondition ws.Cells(12, 9), 2 * aY + 0.5 * bY, 0.1
    
    'Checking abs value of the vectors
    ChangeColorOnCondition ws.Cells(13, 2), AbsVector(aX, aY), 0.1
    ChangeColorOnCondition ws.Cells(13, 3), AbsVector(bX, bY), 0.1
    ChangeColorOnCondition ws.Cells(13, 4), AbsVector(aX + bX, aY + bY), 0.1
    ChangeColorOnCondition ws.Cells(13, 5), AbsVector(aX - bX, aY - bY), 0.1
    ChangeColorOnCondition ws.Cells(13, 6), AbsVector(bX - aX, bY - aY), 0.1
    ChangeColorOnCondition ws.Cells(13, 7), AbsVector(aX + 2 * bX, aY + 2 * bY), 0.1
    ChangeColorOnCondition ws.Cells(13, 8), AbsVector(aX - 2 * bX, aY - 2 * bY), 0.1
    ChangeColorOnCondition ws.Cells(13, 9), AbsVector(2 * aX + 0.5 * bX, 2 * aY + 0.5 * bY), 0.1
    
    'Checking out %error. This needs to be more sofisticated, checking out for the true values, not just what appear in the table.
    For i = 2 To 9
        ChangeColorOnCondition ws.Cells(15, i), PercentError(ws.Cells(14, i), ws.Cells(13, i)), 0.1
        If ws.Cells(15, i).Interior.Color = 204 Then
            ws.Cells(16, i).Value = PercentError(ws.Cells(14, i), ws.Cells(13, i))
        End If
    Next i
    
   
    ChangeColorOnCondition ws.Cells(t2FirstRow, t2FirstCol + 1), aDOTb, 0.01
    ChangeColorOnCondition ws.Cells(t2FirstRow + 2, t2FirstCol + 1), PercentError(ws.Cells(t2FirstRow + 1, t2FirstCol + 1), CDbl(aDOTb)), 0.1
    If ws.Cells(t2FirstRow, t2FirstCol + 1).Interior.ColorIndex = 3 Then
     ws.Cells(t2FirstRow, t2FirstCol + 4).Value = aDOTb
    End If
    If ws.Cells(t2FirstRow + 2, t2FirstCol + 1).Interior.ColorIndex = 3 Then
     ws.Cells(t2FirstRow + 2, t2FirstCol + 4).Value = PercentError(ws.Cells(t2FirstRow + 1, t2FirstCol + 1), CDbl(aDOTb))
    End If
    
    'Checking the angle
    ChangeColorOnCondition ws.Cells(t2FirstRow + 3, t2FirstCol + 1), theta, 0.01
    ChangeColorOnCondition ws.Cells(t2FirstRow + 5, t2FirstCol + 1), PercentError(ws.Cells(t2FirstRow + 4, t2FirstCol + 1), CDbl(theta)), 0.1
    If ws.Cells(t2FirstRow + 3, t2FirstCol + 1).Interior.ColorIndex = 3 Then
     ws.Cells(t2FirstRow + 3, t2FirstCol + 4).Value = theta
    End If
    If ws.Cells(t2FirstRow + 5, t2FirstCol + 1).Interior.ColorIndex = 3 Then
     ws.Cells(t2FirstRow + 5, t2FirstCol + 4).Value = PercentError(ws.Cells(t2FirstRow + 4, t2FirstCol + 1), CDbl(theta))
    End If
    
    'Questions
    For j = 2 To 10 Step 2
        dotProd = 0
        For i = 18 To 20 Step 1
            dotProd = dotProd + ws1.Cells(i, j).Value * ws1.Cells(i, j + 1).Value
        Next i
        If dotProd = 0 Then
            If IsInArray(UCase(ws1.Cells(22, j).Value), negAnswer) Then
                ws1.Cells(22, j).Interior.Color = 204
            End If
            
        Else
            If IsInArray(UCase(ws1.Cells(22, j).Value), posAnswer) Then
                ws1.Cells(22, j).Interior.Color = 204
            End If
        End If
    Next j
End Sub

