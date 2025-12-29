Attribute VB_Name = "AlignArrowAdjustments"
' Align arrow block adjustments to the last selected shape

Sub AlignArrowAdjustments()
    Dim shp As Shape
    Dim lastShape As Shape
    Dim shapeCount As Integer
    Dim i As Integer
    Dim adjustmentCount As Integer
    
    ' 図形が選択されているか確認
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "図形が選択されていません。", vbExclamation
        Exit Sub
    End If
    
    ' 選択された図形の数を確認
    shapeCount = ActiveWindow.Selection.ShapeRange.Count
    If shapeCount < 2 Then
        MsgBox "2つ以上の図形を選択してください。", vbExclamation
        Exit Sub
    End If
    
    ' 最後に選択された図形を取得
    Set lastShape = ActiveWindow.Selection.ShapeRange(shapeCount)
    
    ' 最後に選択された図形が調整可能なパラメータを持っているか確認
    adjustmentCount = lastShape.Adjustments.Count
    If adjustmentCount = 0 Then
        MsgBox "最後に選択された図形は調整可能なパラメータを持っていません。", vbExclamation
        Exit Sub
    End If
    
    ' 他の図形の調整値を変更
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Not shp Is lastShape Then
            If shp.Adjustments.Count = adjustmentCount Then
                ' すべての調整値を基準図形に合わせる
                For i = 1 To adjustmentCount
                    shp.Adjustments(i) = lastShape.Adjustments(i)
                Next i
            End If
        End If
    Next shp
    
    ' クリーンアップ
    Set shp = Nothing
    Set lastShape = Nothing
End Sub
