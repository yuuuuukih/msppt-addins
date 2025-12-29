# MS PowerPoint 自作 AddIns

## 機能を追加する (新たに AddIn を作成する)

1. 機能を追加したい AddIn を選ぶ。新たに作成する場合は template ディレクトリを複製し、適当な名前をつける。
以下では作業ディレクトリ名を `template` として説明し、メインで操作する.pptmファイルも`template.pptm`としているが適宜適切な名前に変更すべきである。

2. 適当な PowerPoint ファイルで VBE を起動し、マクロを作成する。


> [!TIP]
> Visual Basic Editor (VBE) を起動するには、Alt + F11

3. 「プロジェクト」ウィンドウで右クリックし、メニューから「ファイルのエクスポート」を選択して.basファイルで書き出す。
尚、`template/bas` 内に保存すると管理しやすい。

> [!TIP]
> basファイルを保存する際は、Ctrl+Shift+Sからエンコードフォーマットを"ANSI"にしてあげると、メッセージボックスの日本語が文字化けせずに表示される。

5. `template/template.pptm`を開き、VBEを起動する。
その後、「プロジェクト」ウィンドウで右クリックし、メニューから「ファイルのインポート」を選択し、書き出した.basファイルを読み込む。

6. `template/customUI/customUI.xml`を適宜編集する。
以下は例。アイコンは[Office 365アイコン(imageMso)一覧(O)](https://www.ka-net.org/blog/?p=11361)や[ImageMSO List Reference](https://bert-toolkit.com/imagemso-list.html)を参照。
```customUI.xml
<?xml version="1.0" encoding="utf-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon startFromScratch="false">
    <tabs>
      <tab id="MyTab1" label="タブ名" keytip="XX">

        <group id="Group1" label="グループ名1">
          <button id="Button1-1" label="ボタン名1" keytip="A" imageMso="MacroPlay" size="normal" onAction="マクロ名1" />
          <button id="Button1-2" label="ボタン名2" keytip="B" imageMso="MacroPlay" size="normal" onAction="マクロ名2" />
          <button id="Button1-3" label="ボタン名3" keytip="C" imageMso="MacroPlay" size="normal" onAction="マクロ名3" />
        </group>

        <group id="Group2" label="グループ名2">
          <button id="Button2-1" label="ボタン名4" keytip="D" imageMso="MacroPlay" size="normal" onAction="マクロ名4" />
          <button id="Button2-2" label="ボタン名5" keytip="E" imageMso="MacroPlay" size="normal" onAction="マクロ名5" />
          <button id="Button2-3" label="ボタン名6" keytip="F" imageMso="MacroPlay" size="normal" onAction="マクロ名6" />
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>
```

6. `template/template.pptm`を`template/template.zip`に拡張子を変更する。

7.  `customUI.xml`が入った `template/customUI`を、 `template/template.zip`内にコピーする。
また、 `template/template.zip/_rels/.rels`を開き、以下の内容になっていることを確認する。

```.rels
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
<Relationship Id="myCustomUI" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/>
</Relationships>
```

> [!TIP]
> .zip内のファイルは直接編集できないので、編集する必要がある場合は外部にコピーし編集後、適当なディレクトリに移す。

8. `template/template.zip`を`template/template.pptm`に拡張子を戻し、リボンを確認する。

9. `template/template.pptm`を開き、Ctrl+Shift+Sで拡張子を.ppamを選び、`template.ppam`を新たに保存する。
保存先について、Winの場合は、`C:\Users\user_name\AppData\Roaming\Microsoft\AddIns`に置くとよいが必須ではない。

後はこの.ppamファイルを読み込めば自作AddInを使用することができる。
- Win -> 開発>アドイン>PowerPointアドイン>新規追加
- Mac -> ツール>PowerPointアドイン>＋




> [!NOTE]
> [PowerPointでマクロをアドイン化しリボンに追加する方法](https://ppdtp.com/powerpoint/macro-custom-ui/)を参考にしています。

