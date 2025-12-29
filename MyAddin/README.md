# CustomUI

## ObjectSetting

- AlignBlockWidths
  - 最後に選ばれたオブジェクトを基準に、複数のオブジェクトの横幅を揃える。
- AlignBlockHights
  - 最後に選ばれたオブジェクトを基準に、複数のオブジェクトの縦幅を揃える。
- AlignBlockSizes
  - 最後に選ばれたオブジェクトを基準に、複数のオブジェクトの横幅と縦幅を揃える。

## TextSettings

- ToggleTextWrap
  - 選択したオブジェクトにおいて、「図形内でテキストを折り返す」の設定をON/OFFする。
- ClearAllTexts
  - 選択したすべてのオブジェクト内のテキストを削除する。

## MoveObjectsToAlign

- MoveLeftToLastRight
  - 最後に選ばれたオブジェクトを基準とし、基準の右端とその他のオブジェクトの左端が一致するように、その他のオブジェクトを左へ移動する。
- MoveRightToLastLeft
  - 最後に選ばれたオブジェクトを基準とし、基準の左端とその他のオブジェクトの右端が一致するように、その他のオブジェクトを右へ移動する。
- MoveUpToLastBottom
  - 最後に選ばれたオブジェクトを基準とし、基準の下端とその他のオブジェクトの上端が一致するように、その他のオブジェクトを上へ移動する。
- MoveDownToLastTop
  - 最後に選ばれたオブジェクトを基準とし、基準の上端とその他のオブジェクトの下端が一致するように、その他のオブジェクトを下へ移動する。

## ObjectsAlign
- CenterAlignToLast
  - 最後に選択したオブジェクトの位置を動かさずに左右中央揃えをする。
- MiddleAlignToLast
  - 最後に選択したオブジェクトの位置を動かさずに上下中央揃えをする。

## ChangeObjects
- SwapObj
  - オブジェクトの位置を選択した順番に移動させる。例えばA,B,Cの順番にオブジェクトを選択した場合、Bの位置にAを、Cの位置Bを、Aの位置にCを移動する。選択するオブジェクトが2つの場合、位置がSwapするような挙動を示す。
- ReplaceObj
  - 選択したオブジェクト全てを、最後に選択したオブジェクトに置換する。

## InsertTemplate
- InsertCaption
  - キャプションの入力欄を、既定の位置に挿入する。
- InsertSource
  - 出所の入力欄を、規定の位置に挿入する。

## CreateIncircle
三角形を選択した状態で実行すると内接円を作成する。
但し、`ForTriangle`と`ForInvertedTriangle`では、三角形の条件に少し違いがある。
- ForTriangle
  - 選択される三角形は二等辺三角形で底辺は水平であり、かつ上向き三角形（頂点が底辺より上側にある）でなくてはならない。
- ForInvertedTriangle
  - 選択される三角形は二等辺三角形で底辺は水平であり、かつ下向き三角形（頂点が底辺より下側にある）でなくてはならない。

## ShapeSettings
- LockAspectRatio
  - 選択したすべての図形のアスペクト比（縦横比）を固定する。

## ShapeAdjustments
- AlignRoundedCorners
  - 最後に選択した図形の角丸半径を基準に、その他の選択図形の角丸半径を統一する。角丸四角形などの調整可能な図形に対して有効。
- AlignArrowAdjustments
  - 最後に選択した図形の調整値（矢印の角度、五角形の形状など）を基準に、その他の選択図形の調整値を統一する。矢印ブロックや五角形などの調整可能な図形に対して有効。
