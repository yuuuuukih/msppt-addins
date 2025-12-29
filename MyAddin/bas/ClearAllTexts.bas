Attribute VB_Name = "ClearAllTexts"
' Clear all text from selected shapes

Sub ClearAllTexts()
    Dim shp As Shape
    
    ' 図形が選択されているか確認
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "図形が選択されていません。", vbExclamation
        Exit Sub
    End If
    
    ' すべての選択図形のテキストを削除
    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                shp.TextFrame.TextRange.Text = ""
            End If
        End If
    Next shp
    
    ' クリーンアップ
    Set shp = Nothing
End Sub
