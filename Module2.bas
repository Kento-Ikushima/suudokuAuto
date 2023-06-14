Attribute VB_Name = "Module2"
Sub SolveAndVisualize(ByRef board() As Integer)
    Dim row As Integer
    Dim col As Integer
    
    If Not FindUnassignedCell(board, row, col) Then
        ' �������������ꍇ�̓��[�N�V�[�g��Ɍ��ʂ�\������
        ShowSolutionOnWorksheet board
        Exit Sub
    End If
    
    For num = 1 To 9
        If IsSafe(board, row, col, num) Then
            board(row, col) = num
            
            ' ���[�N�V�[�g��Ŏ������ꂽ�l��\������
            ShowAttemptOnWorksheet board, row, col
            
            ' �ċA�I�Ɏ��̃Z���ɐi��
            SolveAndVisualize board
            
            ' �o�b�N�g���b�N�i�����������Ȃ������ꍇ�ɖ߂�j���Ɏ������ꂽ�l���N���A����
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
    Application.Wait (Now + TimeValue("0:00:01")) ' 1�b�ԑҋ@���ĉ���
End Sub


Sub ClearAttemptOnWorksheet(ByRef row As Integer, ByVal col As Integer)
    Cells(row + 1, col + 1).ClearContents
    Application.Wait (Now + TimeValue("0:00:01")) ' 1�b�ԑҋ@���ĉ���
End Sub

Sub VisualizeSudokuSolution()
    Dim board(1 To 9, 1 To 9) As Integer
    Dim i As Integer, j As Integer
    
    ' ���[�N�V�[�g�uSheet1�v��B2����_�ɒl���擾���Aboard�Ɋi�[����
    For i = 1 To 9
        For j = 1 To 9
            board(i, j) = Worksheets("Sheet1").Cells(i + 1, j + 1).Value
        Next j
    Next i
    
    ' board�ɒl���i�[���ꂽ���Ƃ��m�F���邽�߁A�f�o�b�O�v�����g�ŕ\������
    For i = 1 To 9
        For j = 1 To 9
            Debug.Print board(i, j)
        Next j
    Next i
    
    ' ���Ƃ̉������߂�
    SolveAndVisualize board
End Sub

