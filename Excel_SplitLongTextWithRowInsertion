Sub SplitLongTextWithRowInsertion()
    Dim cell As Range
    Dim content As String
    Dim maxLength As Integer
    Dim startPos As Integer
    Dim splitPos As Integer
    Dim line As String
    Dim currentRow As Long
    
    ' 文字数の制限を設定（例：500文字）
    maxLength = 500
    
    ' 選択したセルをループ
    For Each cell In Selection
        content = cell.Value
        startPos = 1
        currentRow = cell.Row
        
        ' 文字数制限を超える場合に分割
        Do While startPos <= Len(content)
            ' 分割位置を決定
            If startPos + maxLength - 1 > Len(content) Then
                splitPos = Len(content)
            Else
                splitPos = startPos + maxLength - 1
            End If
            
            ' 分割した文字列を取得
            line = Mid(content, startPos, splitPos - startPos + 1)
            
            ' 分割した文字列をセルに書き込む
            Cells(currentRow, cell.Column).Value = line
            
            ' 次の開始位置を設定
            startPos = splitPos + 1
            
            ' 次の行に移動
            If startPos <= Len(content) Then
                currentRow = currentRow + 1
                Rows(currentRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            End If
        Loop
    Next cell
End Sub
