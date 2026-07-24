# Project_24 引当・取り置き全シート同期 引継ぎ

更新日: 2026-07-24

## 現在の状態

- ブランチ: `codex/full-allocation-rebuild-v3`
- Project_24へ安全push済み。push直後の `gas_pull_sync.ps1 Project_24` でオンラインとローカルの一致を確認済み。
- Project_24のテスト16本、合計399件が成功。
- Project_24の `.js` / `.gs` 20ファイルが `node --check` 成功。
- 実装仕様: `docs/superpowers/specs/2026-07-24-all-derived-sheets-sync-design.md`
- 実装計画: `docs/superpowers/plans/2026-07-24-all-derived-sheets-sync.md`

## 実装した内容

1. `Project_24/全シート同期.js` に `引当_数値変更後全同期_` を追加。
   - 必要時だけEMS在庫を更新
   - ②引当を実行
   - 取り置き登録、キャンセル戻し確認、Yahoo戻し候補を同じ台帳時点へ更新
   - 失敗時は `引当_整合状態` を削除し、同期失敗として記録・表示

2. 数量を変える最上位処理を中央同期へ接続。
   - 取り置き登録・戻し確定・Yahoo戻し確定・手動解除・孤児解除
   - 個別引当・個別キャンセル
   - 現物確認移行
   - GoQ CSV取込、発送済みCSV取込、消込更新・クリア、注文キャンセル
   - P列書き直し、便の引き直し、EMS更新後引当
   - 全件再計算は管理シート同期失敗でも停止

3. Yahoo CSVを手動運用へ固定。
   - `日本在庫`を確認後、`📤 CSVデータ作成`ボタンで全便＋戻りを出力
   - ⑤便締めからCSVを自動作成しない
   - ⑤は、手動出力記録に対象EMS・商品コード・数量が一致している場合だけ続行
   - CSV作成後に数量が変われば⑤を停止し、再確認・再出力を要求
   - ⑤完了後はEMS在庫・引当・取り置き・分類シートを全同期

## 日常運用

1. 数量や取り置きを更新する（各処理の完了時に全派生シートが自動同期される）。
2. 便を締める前に `日本在庫` を目視確認する。
3. `日本在庫` の `📤 CSVデータ作成` ボタンを押す。
4. 出力をYahooへ反映する。
5. ⑤便締めを実行する。手動出力時点と数量が違えば安全停止する。

既存データを新ロジックで一度揃える場合は、通常の `② 引き当て実行` を1回実行する。⑤便締めはこの初回同期には使わない。

## 重要なリポジトリ運用

- 作業開始時: `git pull` → `powershell -ExecutionPolicy Bypass -File tools\gas_pull_sync.ps1 Project_24`
- 素の `clasp push` は禁止。
- 反映は必ず `powershell -ExecutionPolicy Bypass -File tools\gas_safe_push.ps1 Project_24`
- オンライン編集がある前提で、同期コミットを削除・上書きしない。

## 検証コマンド

```powershell
Get-ChildItem tests\project24_*.test.js | ForEach-Object { node $_.FullName }
Get-ChildItem Project_24 -File | Where-Object { $_.Extension -in '.js','.gs' } | ForEach-Object { node --check $_.FullName }
```

## 次に確認すること

- ユーザーが通常の `② 引き当て実行` を1回行い、実データの全派生シートを現時点へ揃える。
- その後、代表注文で「注文数量 = 確保済み + 不足」と、同一EMS・商品コードの供給上限を実シート上で確認する。
- ⑤の実運用では、CSV未作成時に停止し、CSV作成後は対象便だけ締められることを確認する。
