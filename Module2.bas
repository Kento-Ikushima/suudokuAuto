Attribute VB_Name = "Module2"
Sub SolveAndVisualize(ByRef board() As Integer)
    Dim row As Integer
    Dim col As Integer
    
    If Not FindUnassignedCell(board, row, col) Then
        ' 解が見つかった場合はワークシート上に結果を表示する
        ShowSolutionOnWorksheet board
        Exit Sub
    End If
    
    For num = 1 To 9
        If IsSafe(board, row, col, num) Then
            board(row, col) = num
            
            ' ワークシート上で試し入れた値を表示する
            ShowAttemptOnWorksheet board, row, col
            
            ' 再帰的に次のセルに進む
            SolveAndVisualize board
            
            ' バックトラック（解を見つけられなかった場合に戻る）時に試し入れた値をクリアする
            ClearAttemptOnWorksheet row, col
            
            board(row, col) = 0
        End If
    Next num
End Sub

Sub ShowSolutionOnWorksheet(ByRef board() As Integer)
    For i = 1 To 9
        For j = 1 To 9
            Cells(i + 1, j + 1).Value = board(i, j)
        Next j
    Next i
End Sub

Sub ShowAttemptOnWorksheet(ByRef board() As Integer, ByVal row As Integer, ByVal col As Integer)
    Cells(row + 1, col + 1).Value = board(row, col)
    Application.Wait (Now + TimeValue("0:00:01")) ' 1秒間待機して可視化
End Sub


Sub ClearAttemptOnWorksheet(ByRef row As Integer, ByVal col As Integer)
    Cells(row + 1, col + 1).ClearContents
    Application.Wait (Now + TimeValue("0:00:01")) ' 1秒間待機して可視化
End Sub

Sub VisualizeSudokuSolution()
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
    SolveAndVisualize board
End Sub

