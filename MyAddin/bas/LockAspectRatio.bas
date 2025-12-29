Attribute VB_Name = "LockAspectRatio"
' Toggle aspect ratio lock for selected shapes

Sub LockAspectRatio()
    Dim shp As Shape
    
    ' 図形が選択されているか確認
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "図形が選択されていません。", vbExclamation
        Exit Sub
    End If
    
    ' すべての選択図形のアスペクト比をトグル
    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.LockAspectRatio = msoTrue Then
            ' 固定されている場合は解除
            shp.LockAspectRatio = msoFalse
        Else
            ' 固定されていない場合は固定
            shp.LockAspectRatio = msoTrue
        End If
    Next shp
    
    ' クリーンアップ
    Set shp = Nothing
End Sub
