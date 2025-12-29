Attribute VB_Name = "InsertCaption"

Sub InsertCaption()
    Dim slide As slide
    Dim line As shape
    Dim rect As shape
    Dim groupItems As ShapeRange
    Dim groupShape As shape
    Dim captionText As String
    Dim fontColor As Long
    Dim lineColor As Long
    Dim cmToPoints As Single
    Dim leftPosition As Single
    Dim topPosition As Single
    
    ' cmをポイントに変換する定数
    cmToPoints = 28.3465
    
    ' 現在のアクティブスライドを取得
    Set slide = ActiveWindow.View.slide
    
    ' テキストと色を設定
    captionText = "caption"
    fontColor = RGB(0, 15, 120)
    lineColor = RGB(0, 15, 120)
    
    ' 線を作成
    Set line = slide.Shapes.AddLine(BeginX:=0, BeginY:=0, EndX:=26.4 * cmToPoints, EndY:=0)
    With line.line
        .Weight = 1.5
        .ForeColor.RGB = lineColor
    End With
    
    ' 長方形を作成
    Set rect = slide.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=24.25 * cmToPoints, Height:=1 * cmToPoints)
    With rect
        .Fill.Transparency = 1
        .line.Transparency = 1
        .textFrame.TextRange.Text = captionText
        With .textFrame.TextRange.Font
            .Name = "Yu Gothic UI"
            .Size = 14
            .Bold = msoTrue
            .Color.RGB = fontColor
        End With
        .textFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
    End With
    
    ' 線と長方形を下揃えかつ左揃えで合わせる
    line.Top = rect.Top + rect.Height - line.Height
    line.Left = rect.Left
    
    ' 線と長方形をグループ化する
    Set groupItems = slide.Shapes.Range(Array(line.Name, rect.Name))
    Set groupShape = groupItems.Group
    
    ' グループ化したオブジェクトを指定の位置に配置する
    leftPosition = 0.56 * cmToPoints
    topPosition = 4.29 * cmToPoints
    groupShape.Left = leftPosition
    groupShape.Top = topPosition
    
    ' クリーンアップ
    Set line = Nothing
    Set rect = Nothing
    Set groupItems = Nothing
    Set groupShape = Nothing
    Set slide = Nothing
End Sub

