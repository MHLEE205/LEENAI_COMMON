# LEENAI_COMMON — THOR 作業規則

## 現在このファイルを使っている TOOL
- SI_COMPASS (D-04) v2.0 以降

## 修正規則

1. 修正前に上の「使用 TOOL リスト」を確認し、보고서に記載する。
2. 互換が崩れる変更 (API 名・引数・トークン名変更) は `v2/` 新パスに。`v1/` は修正のみ、削除禁止。
3. ファイル先頭コメントのバージョンと変更履歴を更新。
4. push 前に `node --check v1/leenai-common.js` で文法確認。
5. 問題発生時は即前コミットに revert し、형님에게 報告。

✦ Powered by LEENAI Automation System
