
' 引数パスのファイル内容を返却する
Function ReadFile(path As String)
    Dim buf As String
    Open path For Input As #1
        Do Until EOF(1)
            Line Input #1, buf
        Loop
    Close #1
    ReadFile = buf
End Function

Function RewriteLines(lines() As String)
    ' 書き換えて返却する行データ
    Dim writeLines() As String
    ReDim writeLines(0)
    ' 書き換える行のインデックス
    Dim writeIdx As Integer
    ' 列データ確認中か
    Dim isRetsuCheckMode As Boolean: isRetsuCheckMode = False
    
    ' 書き換え条件
    Dim targetSections() As String: targetSections = Split("[Section1],[Section3]", ",")
    Dim retsuParam() As Variant: retsuParam = Array(2, 1)
    Dim incrementNum As Integer: incrementNum = 500
    Dim targetRetsuIdx As Integer: targetRetsuIdx = 0
    Dim targetRetsu As Integer: targetRetsu = 1 ' セクションごとの列位置をどうするか★
    
    ' 1行ずつループ
    For i = 0 To UBound(lines)
        ' 列ごとに分割する
        Dim retsu() As String: retsu = Split(lines(i), ",")
        ' 列無しの行か
        If UBound(retsu) < 0 Then
            ' 列データ確認中フラグをリセット
            isRetsuCheckMode = False
            ' 空行を書き込む
            Call CreateLineData(writeIdx, writeLines, "")
            ' 次の行確認へ
            GoTo Continue
        End If
        
        If isRetsuCheckMode Or CheckTargetSection(retsu(0), targetSections, targetRetsuIdx) Then
            isRetsuCheckMode = True
            targetRetsu = retsuParam(targetRetsuIdx)
            Dim newbuf As String: newbuf = ""
            ' 列をループ
            For j = 0 To UBound(retsu)
                ' 指定列か
                If j = targetRetsu Then
                    Dim targetNum As Integer: targetNum = Val(retsu(j))
                    targetNum = targetNum + incrementNum
                    retsu(j) = targetNum
                End If
                ' セクション行以外か
                If j <> 0 Then
                    newbuf = newbuf + "," + retsu(j)
                Else
                    newbuf = newbuf + retsu(j)
                End If
            Next j
            Call CreateLineData(UBound(writeLines) + 1, writeLines, newbuf)
        Else
            ' 読み込んだ行そのまま書き込み
            Call CreateLineData(UBound(writeLines) + 1, writeLines, lines(i))
        End If
Continue:
    Next
    
    RewriteLines = writeLines
    
End Function

Function CheckTargetSection(data As String, sections() As String, retsuIdx As Integer)
    Dim isTarget As Boolean: isTarget = False
    For i = 0 To UBound(sections)
        If data = sections(i) Then
            isTarget = True
            retsuIdx = i
            Exit For
        End If
    Next
    CheckTargetSection = isTarget
End Function

Sub CreateLineData(writeIdx As Integer, writeLines() As String, lineData As String)
    ' 行データの最終行＋１を書き込み可能にしてlineDataを書き込む
    writeIdx = UBound(writeLines) + 1
    ReDim Preserve writeLines(writeIdx)
    writeLines(writeIdx) = lineData
End Sub


Sub sample()
    Dim openFilePath As String: openFilePath = ThisWorkbook.path & "\readme.txt"
    Dim writeFilePath As String: writeFilePath = ThisWorkbook.path & "\readme_out.txt"
    
    Debug.Print "---- " & Now() & " ----"
    
    ' 引数ファイルの内容を取得
    Dim contents As String: contents = ReadFile(openFilePath)
    ' 全行を配列に取得
    Dim lines() As String: lines = Split(contents, vbLf)
    
    ' 書き込みデータ取得
    Dim writeLines() As String: writeLines = RewriteLines(lines)
    
    For k = 0 To UBound(writeLines)
        Debug.Print writeLines(k) & vbLf;
    Next
    
    
    Open writeFilePath For Output As #2
        For i = 1 To UBound(writeLines)
            If i <> UBound(writeLines) Then
                Print #2, writeLines(i) & vbLf;
            End If
        Next i
    Close #2
    
End Sub
