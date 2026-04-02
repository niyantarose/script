# adapter spec

- `match(url)` で対象 URL を判定する
- `collectRaw(tabId)` で content script に probe を依頼する
- DOM 解析は adapter 内ではなく content script / rawExtract に寄せる
- downloads API は呼ばない
