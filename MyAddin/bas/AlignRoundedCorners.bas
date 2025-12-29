Attribute VB_Name = "AlignRoundedCorners"
' Align rounded corner radius to the last selected shape

Sub AlignRoundedCorners()
    Dim shp As Shape
    Dim lastShape As Shape
    Dim shapeCount As Integer
    Dim targetRadius As Single
    
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
    
    ' 最後に選択された図形が調整可能な角丸を持っているか確認
    If lastShape.Adjustments.Count = 0 Then
        MsgBox "最後に選択された図形は角丸の調整ができません。", vbExclamation
        Exit Sub
    End If
    
    ' 基準となる角丸の半径を取得
    targetRadius = lastShape.Adjustments(1)
    
    ' 他の図形の角丸半径を調整
    For Each shp In ActiveWindow.Selection.ShapeRange
        If Not shp Is lastShape Then
            If shp.Adjustments.Count > 0 Then
                shp.Adjustments(1) = targetRadius
            End If
        End If
    Next shp
    
    ' クリーンアップ
    Set shp = Nothing
    Set lastShape = Nothing
End Sub
