# NotebookLM 連携セットアップ

実行日: 2026-09-03

## 調べた目的・問い

Claude Code から NotebookLM を操作できるようにする。あわせて、
NotebookLM 経由の回答が**出典付きで返ってくるか**を実機で検証する
（CLAUDE.md ルール1が実際に機能するかの確認）。

## 結論・要点

- `notebooklm-mcp-cli` v0.10.0 でセットアップ完了。認証済みプロファイルは
  `C:\Users\<user>\.notebooklm-mcp-cli\profiles\default` に保存され、次回以降ログイン不要。
- **出典付き回答は取得できる。** `nlm notebook query` の応答には
  `sources_used` / `citations` / `references` が構造化されて含まれ、
  `references[].cited_text` に引用元の実文章が入る。表は `cited_table` として
  列・行構造で取得できる。
- **ただし出典の中身は検証が必要。** 今回の実測で、Wikipedia インフォボックスの
  抽出結果に誤った行が混入していた（後述）。出典を確認せず要約だけを採用すると
  誤情報をそのまま流すことになる。ルール1の「出典なしの回答は採用しない」は
  実務上必須と確認できた。

## 環境構築でつまずいた点と対処

| 事象 | 原因 | 対処 |
|---|---|---|
| `python3 --version` が `Python` とだけ返る | Windows の App Execution Alias（中身のないダミー `python3.exe`） | python.org 版 Python をインストール |
| `py` コマンドが存在しない | Python 自体が未インストール | python.org から 3.13.15 を導入（`Add python.exe to PATH` を有効化） |
| （参考）Linux 環境で `pip install` が失敗 | ディストリ管理下の PyJWT と衝突（`RECORD file not found`） | `uv tool install` を使用 |

Windows では `python3` ではなく **`py`** を使う。`nlm login` は Chrome を
CDP（Chrome DevTools Protocol）経由で起動するため、**実行前に Chrome の
完全終了（常駐プロセス含む）が必須**。

## 検証手順

```
nlm notebook create "セットアップテスト"
  -> notebook_id: 4044acda-4bd1-4a52-b362-ed8daf6193c5
nlm source add <id> --url https://ja.wikipedia.org/wiki/Google_Apps_Script --wait
  -> source_id: ef07a558-5e63-43dc-b76c-ef5b80a266dd
nlm notebook query <id> "Google Apps Scriptとは何か... 必ず出典を明示して3点で"
```

## 根拠と出典

ソース: 日本語版 Wikipedia「Google Apps Script」
（notebook「セットアップテスト」/ source_id `ef07a558-5e63-43dc-b76c-ef5b80a266dd`）

- **citation [2]** — GAS は Google Workspace 上の軽量アプリケーション開発
  プラットフォームであり、主に Google のサービスを自動化するスクリプト言語。
  JavaScript がもとになっており、開発環境は Google Chrome だけでよい。
  当初 Mike Harm が Google スプレッドシート開発者として働いていた際の
  サイドプロジェクトとして開発。2009年5月に Jonathan Rochelle が
  ベータテストプログラムを発表し、2009年8月に初公開。
- **citation [3]** — 2020年3月、従来の Rhino インタープリタに加えて V8 の提供を
  開始し、ECMAScript6 以降の機能に対応。Apps Script プロジェクトは Google の
  インフラストラクチャでサーバー側で実行される。
- **citation [4]** — 主にスプレッドシートで使われるが適応範囲は広く、
  「SNS に返信が届いた場合 Gmail に届ける」「Gmail の期限の近いタスクを
  Slack で表示する」といった用途も可能。
- **citation [5]** — 2014年3月、ドキュメントとスプレッドシートへアドオンを導入。
  まもなくフォームにも導入された。
- **citation [1]（`cited_table`）** — インフォボックス。開発元 Google、
  種別 マクロ言語、影響を受けた言語 JavaScript。
  **⚠ このテーブル抽出に `['プラットフォーム', 'Instagram']` という誤った行が
  混入していた。** 原文インフォボックスの解析ノイズと判断され、事実として
  採用してはならない。出典確認の必要性を示す実例。

## 残作業

- `nlm setup add claude-code`（Claude Code 本体との MCP 接続）
