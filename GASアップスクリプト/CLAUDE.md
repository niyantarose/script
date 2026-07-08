# CLAUDE.md

このリポジトリの作業ルールは [AGENTS.md](AGENTS.md) を参照してください。特に：

- **どこかで編集されている前提で動く**：編集前に必ず `git pull` →
  `tools/gas_pull_sync.ps1 Project_XX` でオンライン編集を取り込む
- **Project_19への反映は素のclasp push禁止**・必ず `tools/p19_safe_push.ps1` を使う
