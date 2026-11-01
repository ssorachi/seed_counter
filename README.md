# 種子カウンター

トレイで撮ったシャーレの写真から種子を数え、封筒のバーコードを読み取って、結果をExcelに1枚1行で記録します。写真のコピーはバーコード（対応表があれば系統名）の名前で保存し、輪郭を描いた確認画像も保存します。

## 使い方

1. ［写真を追加］で、トレイを撮った写真を選びます。HEICも使えます。
2. ［開始］を押すと、写真ごとに種子を数え、バーコードを読み取ります。
3. 黄色の行は確認が必要な写真です。［要確認の写真を入力］で、ラベルを見ながらバーコードを入力します。
4. 数え方が合わないときは［しきい値の調整］で設定を変え、右クリックの［数え直す］で処理済みの写真に反映します。

主なキー操作：Ctrl+O 写真を追加、F5 開始、Esc 中止、Enter バーコードを入力、Ctrl+R 数え直す、Delete リストから外す（Macは Ctrl の代わりに Cmd）。

## 系統名の対応表

［ツール］メニューの［対応表を開く］でExcel（.xlsx）かCSVを読み込むと、写真の名前とExcelに系統名を使います。

- 形式は `mapping_template.xlsx`（対応表の見本） のとおりです。1列目にバーコード、2列目に系統名を入れます。見出しは無くてもかまいません。
- ［対応表の見本を保存］で、処理済みの写真のバーコードが入った表を作れます。
- 対応表にないバーコードは要確認になります。入力画面では系統名を入力してもかまいません。
- バーコードが数字だけの場合は、先頭の0が消えないよう、その列を文字列の書式にしてください。

## ダウンロード（ビルド済み）

Releases、またはActionsの最新のbuildにあるArtifactsから取得します。

- `SeedCounter-Windows.zip`：展開して `SeedCounter\SeedCounter.exe` を起動します。
- `SeedCounter-Mac-AppleSilicon.zip`、`SeedCounter-Mac-Intel.zip`：展開して `SeedCounter.app` を起動します。
- Macで「開発元を確認できない」と表示されたら、右クリックして［開く］を選びます。ターミナルで `xattr -dr com.apple.quarantine SeedCounter.app` を実行してもかまいません。
- WindowsでSmartScreenが表示されたら、［詳細情報］を押してから［実行］を押します。

## ビルドのしかた

- `main` にpushすると自動でビルドされます。Actions画面の［Run workflow］からも実行できます。
- `v1.0` のようなタグをpushすると、Releasesにも添付されます。
  ```
  git tag v1.0 && git push origin v1.0
  ```
- `tests/` に写真（.jpg、.heic）を入れておくと、ビルドしたアプリでその写真を処理して動作を確認します。

## Pythonで直接動かす場合

```
pip install -r requirements.txt
python seed_counter_app.py
```
