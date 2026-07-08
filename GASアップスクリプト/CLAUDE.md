# CLAUDE.md

このリポジトリの作業ルールは [AGENTS.md](AGENTS.md) を参照してください。特に：

- **どこかで編集されている前提で動く**：編集前に必ず `git pull` →
  `tools/gas_pull_sync.ps1 Project_XX` でオンライン編集を取り込む
- **素のclasp pushは全プロジェクト（01〜24）で禁止**・反映は必ず
  `tools/gas_safe_push.ps1 Project_XX`（P19は `tools/p19_safe_push.ps1` でも同じ）
