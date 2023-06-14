Attribute VB_Name = "Module1"
Function FindUnassignedCell(ByRef board() As Integer, ByRef row As Integer, ByRef col As Integer) As Boolean
    For row = 1 To 9
        For col = 1 To 9
            If board(row, col) = 0 Then
                FindUnassignedCell = True
                Exit Function
            End If
        Next col
    Next row
    
    FindUnassignedCell = False
End Function

Function Solve(ByRef board() As Integer) As Boolean
    Dim row As Integer
    Dim col As Integer
    
    If Not FindUnassignedCell(board, row, col) Then
        ' 未割り当てのセルがない場合は解が見つかったとする
        Solve = True
        Exit Function
    End If
    
    For num = 1 To 9
        If IsSafe(board, row, col, num) Then
            board(row, col) = num
            
            If Solve(board) Then
                Solve = True
                Exit Function
            End If
            
            board(row, col) = 0 ' バックトラック（解を見つけられなかった場合に戻る）
        End If
    Next num
    
    Solve = False
End Function

Function IsSafe(ByRef board() As Integer, ByVal row As Integer, ByVal col As Integer, ByVal num As Integer) As Boolean
    ' 同じ行に同じ数字が存在しないかチェック
    For c = 1 To 9
        If board(row, c) = num Then
            IsSafe = False
            Exit Function
        End If
    Next c
    
    ' 同じ列に同じ数字が存在しないかチェック
    For r = 1 To 9
        If board(r, col) = num Then
            IsSafe = False
            Exit Function
        End If
    Next r
    
    ' 同じ3x3のボックスに同じ数字が存在しないかチェック
    Dim startRow As Integer
    Dim startCol As Integer
    startRow = 3 * Int((row - 1) / 3) + 1
    startCol = 3 * Int((col - 1) / 3) + 1
    
    For r = startRow To startRow + 2
        For c = startCol To startCol + 2
            If board(r, c) = num Then
                IsSafe = False
                Exit Function
            End If
        Next c
    Next r
    
    IsSafe = True
End Function

Sub SolveSudoku()
    Dim board(1 To 9, 1 To 9) As Integer
    Dim i As Integer, j As Integer
    
    ' ワークシート「Sheet1」のB2を基点に値を取得し、boardに格納する
    For i = 1 To 9
        For j = 1 To 9
            board(i, j) = Worksheets("Sheet1").Cells(i + 1, j + 1).Value
        Next j
    Next i
    
    ' boardに値が格納されたことを確認するため、デバッグプリントで表示する
    For i = 1 To 9
        For j = 1 To 9
            Debug.Print board(i, j)
        Next j
    Next i
    
    ' 数独の解を求める
    If Solve(board) Then
        ' 解を表示
        For i = 1 To 9
            For j = 1 To 9
                Cells(i + 1, j + 1).Value = board(i, j)
            Next j
        Next i
    Else
        MsgBox "解が見つかりませんでした。"
    End If
End Sub

