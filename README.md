# LEENAI_COMMON

LEENAI 全 TOOL が読み込む共通 CSS・JS ファイル (DOC-Z-05 v1.0)。

## 使い方

```html
<link rel="stylesheet" href="https://mhlee205.github.io/LEENAI_COMMON/v1/leenai-common.css">
<script src="https://mhlee205.github.io/LEENAI_COMMON/v1/leenai-common.js"></script>
<script>
LEENAI.init({
  sys:     'D-04',
  name:    '保険入力支援',
  version: '2.0',
  date:    '2026-10-07',
  desc:    'COMPASS と LC JSON から保険申込の入力値を作成します',
  onReady: startTool
});
</script>
```

## ロード失敗時のフォールバック (TOOL 側に必ず入れること)

```html
<script>
if(!window.LEENAI){document.addEventListener('DOMContentLoaded',function(){document.body.innerHTML=
'<div style="padding:40px;font-family:sans-serif">共通ファイルの読み込みに失敗しました。ページを再読み込みしてください。</div>';});}
</script>
```

## 公開 API

| API | 説明 |
|---|---|
| `LEENAI.init(cfg)` | 初期化・ログイン画面表示。cfg: sys/name/version/date/desc/onReady/needGraph |
| `await LEENAI.auth.token('dataverse')` | DV トークン (期限切れで再ログイン誘導) |
| `await LEENAI.auth.token('graph')` | Graph トークン (失敗時 null) |
| `LEENAI.auth.status` | `{dataverse, graph, graphError}` |
| `LEENAI.user` | `{name, email, staffCode, initials}` |
| `LEENAI.logout()` | ログアウト → 第一画面 |
| `LEENAI.loading.show(msg)/.update(msg)/.hide()` | ローディング |
| `LEENAI.toast(msg, type)` | トースト (ok/warn/danger/info) |
| `LEENAI.addHeaderButton({label,onClick,title})` | ヘッダーボタン追加 |
| `LEENAI.STAFF` | 社員マスタ 7 名 |
| `LEENAI.COMPANY` | 会社 TEL/FAX |
| `LEENAI.VERSION` | `'v1.0'` |

## バージョン管理

- `v1/` は破壊的変更禁止。互換が崩れる変更は `v2/` 以降。
- 修正前にこのファイルを使う TOOL 一覧を確認し、CLAUDE.md の規則に従う。

✦ Powered by LEENAI Automation System
