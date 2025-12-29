Attribute VB_Name = "DecreaseAndDistribute"
' オブジェクトのサイズを小さくしながら均等配置する

Sub DecreaseAndDistribute()
    Call AdjustSizeAndDistribute(-0.1)
End Sub

' オブジェクトのサイズを調整して均等配置する
' sizeChange: サイズの変更量（cm単位、正の値で拡大、負の値で縮小）
Sub AdjustSizeAndDistribute(sizeChange As Single)
    Dim shapeRange As shapeRange
    Dim shapeCount As Integer
    Dim i As Integer
    Dim tolerance As Single
    Dim isHorizontalMode As Boolean
    Dim isVerticalMode As Boolean
    
    ' 許容誤差（ポイント単位）
    tolerance = 1
    
    ' オブジェクトが選択されているか確認
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "オブジェクトが選択されていません。", vbExclamation
        Exit Sub
    End If
    
    Set shapeRange = ActiveWindow.Selection.shapeRange
    shapeCount = shapeRange.Count
    
    ' 2つ以上のオブジェクトが選択されているか確認
    If shapeCount < 2 Then
        MsgBox "2つ以上のオブジェクトを選択してください。", vbExclamation
        Exit Sub
    End If
    
    ' 配置方向を判定
    isHorizontalMode = CheckHorizontalAlignment(shapeRange, tolerance)
    isVerticalMode = CheckVerticalAlignment(shapeRange, tolerance)
    
    ' モードの判定
    If isHorizontalMode And Not isVerticalMode Then
        ' 横方向に揃えるモード
        Call AdjustHorizontal(shapeRange, sizeChange)
    ElseIf isVerticalMode And Not isHorizontalMode Then
        ' 縦方向に揃えるモード
        Call AdjustVertical(shapeRange, sizeChange)
    Else
        MsgBox "オブジェクトの配置を検出できませんでした。" & vbCrLf & _
               "横方向に揃える場合：上端または下端を揃えてください。" & vbCrLf & _
               "縦方向に揃える場合：左端または右端を揃えてください。", vbExclamation
        Exit Sub
    End If
    
    ' クリーンアップ
    Set shapeRange = Nothing
End Sub

' 横方向の整列をチェック（上端または下端が揃っているか）
Function CheckHorizontalAlignment(shapeRange As shapeRange, tolerance As Single) As Boolean
    Dim i As Integer
    Dim topAligned As Boolean
    Dim bottomAligned As Boolean
    Dim refTop As Single
    Dim refBottom As Single
    
    topAligned = True
    bottomAligned = True
    
    refTop = shapeRange(1).Top
    refBottom = shapeRange(1).Top + shapeRange(1).Height
    
    For i = 2 To shapeRange.Count
        ' 上端のチェック
        If Abs(shapeRange(i).Top - refTop) > tolerance Then
            topAligned = False
        End If
        
        ' 下端のチェック
        If Abs((shapeRange(i).Top + shapeRange(i).Height) - refBottom) > tolerance Then
            bottomAligned = False
        End If
    Next i
    
    CheckHorizontalAlignment = topAligned Or bottomAligned
End Function

' 縦方向の整列をチェック（左端または右端が揃っているか）
Function CheckVerticalAlignment(shapeRange As shapeRange, tolerance As Single) As Boolean
    Dim i As Integer
    Dim leftAligned As Boolean
    Dim rightAligned As Boolean
    Dim refLeft As Single
    Dim refRight As Single
    
    leftAligned = True
    rightAligned = True
    
    refLeft = shapeRange(1).Left
    refRight = shapeRange(1).Left + shapeRange(1).Width
    
    For i = 2 To shapeRange.Count
        ' 左端のチェック
        If Abs(shapeRange(i).Left - refLeft) > tolerance Then
            leftAligned = False
        End If
        
        ' 右端のチェック
        If Abs((shapeRange(i).Left + shapeRange(i).Width) - refRight) > tolerance Then
            rightAligned = False
        End If
    Next i
    
    CheckVerticalAlignment = leftAligned Or rightAligned
End Function

' 横方向の調整（幅を変更して横に均等配置）
Sub AdjustHorizontal(shapeRange As shapeRange, sizeChange As Single)
    Dim i As Integer
    Dim leftMost As Single
    Dim rightMost As Single
    Dim leftMostIndex As Integer
    Dim rightMostIndex As Integer
    Dim totalWidth As Single
    Dim spacing As Single
    Dim currentLeft As Single
    Dim positions() As Single
    Dim indices() As Integer
    Dim tempPos As Single
    Dim tempIndex As Integer
    Dim j As Integer
    
    ' cm単位をポイント単位に変換（1cm = 28.3465ポイント）
    Dim sizeChangePoints As Single
    sizeChangePoints = sizeChange * 28.3465
    
    ' 最も左と最も右のオブジェクトを特定
    leftMost = shapeRange(1).Left
    rightMost = shapeRange(1).Left + shapeRange(1).Width
    leftMostIndex = 1
    rightMostIndex = 1
    
    For i = 2 To shapeRange.Count
        If shapeRange(i).Left < leftMost Then
            leftMost = shapeRange(i).Left
            leftMostIndex = i
        End If
        If shapeRange(i).Left + shapeRange(i).Width > rightMost Then
            rightMost = shapeRange(i).Left + shapeRange(i).Width
            rightMostIndex = i
        End If
    Next i
    
    ' すべてのオブジェクトの幅を変更
    For i = 1 To shapeRange.Count
        shapeRange(i).Width = shapeRange(i).Width + sizeChangePoints
    Next i
    
    ' オブジェクトを左端の位置でソート
    ReDim positions(1 To shapeRange.Count)
    ReDim indices(1 To shapeRange.Count)
    
    For i = 1 To shapeRange.Count
        positions(i) = shapeRange(i).Left
        indices(i) = i
    Next i
    
    ' バブルソート
    For i = 1 To shapeRange.Count - 1
        For j = i + 1 To shapeRange.Count
            If positions(i) > positions(j) Then
                tempPos = positions(i)
                positions(i) = positions(j)
                positions(j) = tempPos
                
                tempIndex = indices(i)
                indices(i) = indices(j)
                indices(j) = tempIndex
            End If
        Next j
    Next i
    
    ' 総幅を計算
    totalWidth = 0
    For i = 1 To shapeRange.Count
        totalWidth = totalWidth + shapeRange(i).Width
    Next i
    
    ' 間隔を計算
    spacing = (rightMost - leftMost - totalWidth) / (shapeRange.Count - 1)
    
    ' 再配置
    currentLeft = leftMost
    For i = 1 To shapeRange.Count
        shapeRange(indices(i)).Left = currentLeft
        currentLeft = currentLeft + shapeRange(indices(i)).Width + spacing
    Next i
End Sub

' 縦方向の調整（高さを変更して縦に均等配置）
Sub AdjustVertical(shapeRange As shapeRange, sizeChange As Single)
    Dim i As Integer
    Dim topMost As Single
    Dim bottomMost As Single
    Dim topMostIndex As Integer
    Dim bottomMostIndex As Integer
    Dim totalHeight As Single
    Dim spacing As Single
    Dim currentTop As Single
    Dim positions() As Single
    Dim indices() As Integer
    Dim tempPos As Single
    Dim tempIndex As Integer
    Dim j As Integer
    
    ' cm単位をポイント単位に変換
    Dim sizeChangePoints As Single
    sizeChangePoints = sizeChange * 28.3465
    
    ' 最も上と最も下のオブジェクトを特定
    topMost = shapeRange(1).Top
    bottomMost = shapeRange(1).Top + shapeRange(1).Height
    topMostIndex = 1
    bottomMostIndex = 1
    
    For i = 2 To shapeRange.Count
        If shapeRange(i).Top < topMost Then
            topMost = shapeRange(i).Top
            topMostIndex = i
        End If
        If shapeRange(i).Top + shapeRange(i).Height > bottomMost Then
            bottomMost = shapeRange(i).Top + shapeRange(i).Height
            bottomMostIndex = i
        End If
    Next i
    
    ' すべてのオブジェクトの高さを変更
    For i = 1 To shapeRange.Count
        shapeRange(i).Height = shapeRange(i).Height + sizeChangePoints
    Next i
    
    ' オブジェクトを上端の位置でソート
    ReDim positions(1 To shapeRange.Count)
    ReDim indices(1 To shapeRange.Count)
    
    For i = 1 To shapeRange.Count
        positions(i) = shapeRange(i).Top
        indices(i) = i
    Next i
    
    ' バブルソート
    For i = 1 To shapeRange.Count - 1
        For j = i + 1 To shapeRange.Count
            If positions(i) > positions(j) Then
                tempPos = positions(i)
                positions(i) = positions(j)
                positions(j) = tempPos
                
                tempIndex = indices(i)
                indices(i) = indices(j)
                indices(j) = tempIndex
            End If
        Next j
    Next i
    
    ' 総高さを計算
    totalHeight = 0
    For i = 1 To shapeRange.Count
        totalHeight = totalHeight + shapeRange(i).Height
    Next i
    
    ' 間隔を計算
    spacing = (bottomMost - topMost - totalHeight) / (shapeRange.Count - 1)
    
    ' 再配置
    currentTop = topMost
    For i = 1 To shapeRange.Count
        shapeRange(indices(i)).Top = currentTop
        currentTop = currentTop + shapeRange(indices(i)).Height + spacing
    Next i
End Sub
