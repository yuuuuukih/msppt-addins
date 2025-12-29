Attribute VB_Name = "LockAspectRatio"
' Lock aspect ratio for selected shapes

Sub LockAspectRatio()
    Dim shp As Shape
    Dim shapeCount As Integer
    
    ' 図形が選択されているか確認
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "図形が選択されていません。"
        Exit Sub
    End If
    
    ' 選択された図形の数を取得
    shapeCount = ActiveWindow.Selection.ShapeRange.Count
    
    ' すべての選択図形のアスペクト比を固定
    For Each shp In ActiveWindow.Selection.ShapeRange
        shp.LockAspectRatio = msoTrue
    Next shp
    
    MsgBox shapeCount & " 個の図形のアスペクト比を固定しました。"
    
    ' クリーンアップ
    Set shp = Nothing
End Sub
