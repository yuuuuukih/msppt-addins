Attribute VB_Name = "MoveDownToLastTop"
' Move down to align with the top edge of the last selected object

Sub MoveDownToLastTop()
    Dim selectedItem As Shape
    Dim referenceItem As Shape
    Dim slide As slide
    Dim refTopPosition As Single
    Dim i As Integer
    
    ' 現在のアクティブスライドを取得
    Set slide = ActiveWindow.View.slide
    
    ' 選択オブジェクトがない場合は終了
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "オブジェクトが選択されていません。"
        Exit Sub
    End If
    
    ' 最後に選択されたオブジェクトを基準オブジェクトとして設定
    Set referenceItem = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count)
    
    ' 基準オブジェクトの上端の位置を取得
    refTopPosition = referenceItem.Top
    
    ' 他のすべてのオブジェクトを移動
    For i = 1 To ActiveWindow.Selection.ShapeRange.Count - 1
        Set selectedItem = ActiveWindow.Selection.ShapeRange(i)
        
        ' 基準オブジェクト以外を下に移動
        selectedItem.Top = refTopPosition - selectedItem.Height
    Next i
    
    ' クリーンアップ
    Set selectedItem = Nothing
    Set referenceItem = Nothing
    Set slide = Nothing
End Sub

