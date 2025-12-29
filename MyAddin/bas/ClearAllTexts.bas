Attribute VB_Name = "ClearAllTexts"
' Clear all text from selected shapes

Sub ClearAllTexts()
    Dim shp As Shape
    Dim shapeCount As Integer
    Dim clearedCount As Integer
    
    ' 図形が選択されているか確認
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "図形が選択されていません。"
        Exit Sub
    End If
    
    ' 選択された図形の数を取得
    shapeCount = ActiveWindow.Selection.ShapeRange.Count
    clearedCount = 0
    
    ' すべての選択図形のテキストを削除
    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                shp.TextFrame.TextRange.Text = ""
                clearedCount = clearedCount + 1
            End If
        End If
    Next shp
    
    MsgBox clearedCount & " 個の図形からテキストを削除しました。"
    
    ' クリーンアップ
    Set shp = Nothing
End Sub
