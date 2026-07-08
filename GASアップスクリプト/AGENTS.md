# このリポジトリで作業するAIエージェント（Cursor / Claude Code等）への指示

## 体制

- 韓国の共同作業者: claude.aiチャットで作ったコードを **Apps Scriptエディタに直接貼り付け**て更新する
  （主にProject_19の `インボイス.js`、`P-touch.gs.js`、`P-touch-CSV‗UTF16.js`）。GitHub・エディタ・CLIは使わない
- 日本側（リポジトリ管理者）: 職場PC（Claude Code）と自宅PC（Cursor）の2台で作業する

## 大原則: どこかで編集されている前提で動く — 最重要

スクリプトはオンライン側（Apps Scriptエディタへの貼り付け・編集）で**いつでも変更されうる**。
さらにバックグラウンドの自動ツールが、デプロイ版の内容をこのリポジトリの作業ツリーに
**未コミット変更として映してくる**ことがある。ローカルが最新という前提で編集してはいけない。

**編集を始める前に必ず：**

1. `git pull`
2. 対象プロジェクトのオンライン編集を取り込む：
   `powershell -ExecutionPolicy Bypass -File tools\gas_pull_sync.ps1 Project_XX`
   （差分があれば自動で `sync(Project_XX): オンライン編集分を取り込み` コミットが作られる）
3. 作業ツリーに身に覚えのない変更が残っていたら、それはオンライン編集の反映なので
   削除・上書きせず sync コミットとして取り込んでから作業する

**反映（push）するとき — 全プロジェクト共通：**

- **素の `clasp push` は全プロジェクト（Project_01〜24）で禁止。** 必ず安全pushを使う：
  `powershell -ExecutionPolicy Bypass -File tools\gas_safe_push.ps1 Project_XX`
  （Project_19は従来どおり `tools\p19_safe_push.ps1` でも同じ＝中身は共通版への委譲）
- 安全pushは前回push時の基準（tools/sync_state/<Project>.json）との3方向同期で
  全ファイルを保護する: オンラインだけ変更→自動取り込み / ローカルだけ変更→push /
  両方変更→対話確認。基準ファイルはスクリプトが自動更新・コミットする
- 初回実行時にオンライン側へ未知の変更があると対話確認が出ることがある
  （先に `gas_pull_sync.ps1` で取り込んでおけば出ない）

## Project_19（発注EMSリスト）の反映ルール — 最重要

1. **素の `clasp push` は禁止。** オンライン全体をローカル内容で上書きするため、
   共同作業者の貼り付け更新が巻き戻って消える事故が実際に起きた（2026-07-07）
2. 反映は必ず `tools/p19_safe_push.ps1` を使う:
   `powershell -ExecutionPolicy Bypass -File tools\p19_safe_push.ps1`
   （中身は全プロジェクト共通の `gas_safe_push.ps1`。基準は tools/sync_state/Project_19.json）
3. 作業開始時は必ず `git pull`（2台のPCで作業しているため、ローカルが古いことがある）
4. `インボイス.js`・`P-touch系` は主に共同作業者が更新するファイル。こちらから書き換える
   場合は「両方変更」の衝突を避けるため、共同作業者が作業していない時間帯に行い、
   変更後に「上書きした」と連絡すること

## 発注NOの制約（Project_19）

- 発注NO＝「発注日8桁_カート番号2桁_行番号」で全行ユニーク。
  EMSリストの照合キー（購入No-商品コード）がこれに依存する
- **一度付いた番号は変更しない**（EMS大邱・EMSリスト・発注シートが参照している）
- カート番号は日付内max+1で採番し欠番は再利用しない
- 分割発送＝同じ照合キーが複数回出るのは正常。重複判定は数量ベース
- 詳細仕様: `docs/superpowers/specs/2026-07-07-orderno-auto-numbering-v2-design.md`
