# zaiko-tool — user systemd timers

本番では **ubuntu ユーザー** の user systemd に配置し、`systemctl --user` で管理します。

## ログアウト後も動かす（必須）

user タイマーは **linger 無効**だとセッション終了で止まることがあります。

```bash
loginctl show-user ubuntu -p Linger
# Linger=no の場合:
sudo loginctl enable-linger ubuntu
loginctl show-user ubuntu -p Linger   # Linger=yes を確認
```

## 配置パス（本番例）

- アプリ: `/home/ubuntu/zaiko-tool/app`
- ユニット: `~/.config/systemd/user/zaiko-*.service` / `zaiko-*.timer`
- ジョブスクリプト: `/home/ubuntu/zaiko-tool/app/scripts/job_*.sh`

```bash
cp zaiko-*.service zaiko-*.timer ~/.config/systemd/user/
chmod +x /home/ubuntu/zaiko-tool/app/scripts/job_*.sh
systemctl --user daemon-reload
systemctl --user enable --now zaiko-yahoo-orders.timer zaiko-yahoo-stock-diff.timer zaiko-yahoo-stock-full.timer zaiko-run-all-imports.timer
```

## `list-timers` の NEXT が `-` について

`systemctl --user list-timers` の表示が `-` でも、次回時刻は次で確認できます。

```bash
systemctl --user show zaiko-yahoo-orders.timer -p NextElapseUSecRealtime
```

## ロールバック（タイマー停止）

```bash
systemctl --user disable --now zaiko-yahoo-orders.timer zaiko-yahoo-stock-diff.timer zaiko-yahoo-stock-full.timer zaiko-run-all-imports.timer
```
