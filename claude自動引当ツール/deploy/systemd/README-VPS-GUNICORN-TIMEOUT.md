# 本番 VPS: Gunicorn `--timeout` 延長の反映手順

## 前提（重要）

- **`sudo` で `/etc/systemd/system/zaiko-tool.service` を上書きするまで、本番の unit は変更されません。** リポジトリやホーム配下に置いたファイルだけでは、稼働中の Gunicorn の設定は変わりません。
- Cursor からの SSH では `sudo` が非対話で使えない場合があるため、**VPS にログインしたシェルで** 以下を実行してください。

## 目的

- HTTP インポート（`/import/yahoo_orders` など）が長時間かかる場合に、Gunicorn の **WORKER TIMEOUT（既定 120 秒）** で落ちないよう、`--timeout 900` に合わせる。
- ジョブ側の `curl --max-time` を Gunicorn より長めに保つ（orders: **900 秒**、stock full: **1800 秒**）。

## ホーム配下に置くファイル（配置済みを想定）

| ローカル（リポジトリ） | VPS 上の例 |
|------------------------|------------|
| `deploy/systemd/zaiko-tool.service` | `/home/ubuntu/zaiko-tool.service.timeout900` |
| `scripts/job_yahoo_orders.sh` | `/home/ubuntu/zaiko-tool/app/scripts/job_yahoo_orders.sh` |
| `scripts/job_yahoo_stock_full.sh` | `/home/ubuntu/zaiko-tool/app/scripts/job_yahoo_stock_full.sh` |

ジョブスクリプトは **実行権限** が必要です（`chmod +x`）。

```bash
chmod +x /home/ubuntu/zaiko-tool/app/scripts/job_yahoo_orders.sh \
         /home/ubuntu/zaiko-tool/app/scripts/job_yahoo_stock_full.sh
```

## 適用前の確認（差分）

`/etc` の現行 unit と、ホーム配下の候補を比較します。

```bash
diff -u /etc/systemd/system/zaiko-tool.service /home/ubuntu/zaiko-tool.service.timeout900
```

主な差分は `ExecStart` 内の **`--timeout 120` → `--timeout 900`** であることを確認してください。

## バックアップ（推奨）

```bash
sudo cp -a /etc/systemd/system/zaiko-tool.service \
  "/etc/systemd/system/zaiko-tool.service.bak.$(date +%Y%m%d%H%M%S)"
```

## システム unit の反映（要 sudo）

```bash
sudo cp /home/ubuntu/zaiko-tool.service.timeout900 /etc/systemd/system/zaiko-tool.service
sudo systemctl daemon-reload
sudo systemctl restart zaiko-tool.service
sudo systemctl status zaiko-tool.service --no-pager
grep -E 'ExecStart|timeout' /etc/systemd/system/zaiko-tool.service
```

期待: `ExecStart=... gunicorn ... --timeout 900 ...`

## user systemd（タイマー）について

- **ジョブ用の `.sh` だけを差し替えた場合**、`systemctl --user daemon-reload` は **通常不要**です（次回タイマー起動から新しいスクリプトが使われます）。
- **`~/.config/systemd/user/` 内の `.service` / `.timer` を編集した場合**は、反映に必須です。

```bash
systemctl --user daemon-reload
```

## ロールバック（必要時）

バックアップしたファイルを `sudo cp` で戻し、`sudo systemctl daemon-reload` と `sudo systemctl restart zaiko-tool.service` を実行してください。

## 検証のヒント

```bash
journalctl -u zaiko-tool.service -n 80 --no-pager
```

長時間インポート後も **WORKER TIMEOUT** が出ないことを確認します。
