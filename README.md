# What is Ariawase?

Ariawase is an open source VBA library.

## Quick Start

Run `build.bat`, you'll get office macro-enabled file in bin directory.

## 使い方 (vbac.wsf)

Windows のコマンドプロンプトで実行します（`cscript.exe` が必要）。Office アプリは実行前に閉じ、[VBA プロジェクト オブジェクト モデルへの信頼] を有効にしてください。

開発前の準備（自動化）
- PowerShell で環境準備（Office を閉じて信頼を有効化）:
  - `powershell -ExecutionPolicy Bypass -File scripts/prepare-office.ps1`
  - 変更内容の確認のみ: `powershell -File scripts/prepare-office.ps1 -WhatIf`

- 基本構文: `cscript //nologo vbac.wsf <command> [options]`

コマンド
- `combine`: `src/` のコードを `bin/` の Office ファイルへ取り込み（ビルド）。`bak/` にバックアップ作成。
- `decombine`: `bin/` の Office ファイルから VBA を `src/` へ書き出し（エクスポート）。
- `clear`: `bin/` の VBA コンポーネントを削除（ドキュメントモジュールの宣言のみ残す）。
- `help`: ヘルプを表示。

主なオプション
- `/binary:<dir>`: バイナリ（Office ファイル）ディレクトリ（既定: `bin`）
- `/source:<dir>`: ソースコードディレクトリ（既定: `src`）
- `/vbaproj`: `src/App.vbaproj` を用いて参照設定なども反映
- `/dbcompact`: Access データベースをコンパクト化（Access のみ）

使用例
- エクスポート（編集用に `src/` を更新）:
  - `cscript //nologo vbac.wsf decombine /binary:bin /source:src`
- インポート（ビルドして `bin/` を更新）:
  - `cscript //nologo vbac.wsf combine /source:src /binary:bin /vbaproj`
- Access のコンパクト化込みでビルド:
  - `cscript //nologo vbac.wsf combine /dbcompact`

推奨ワークフロー
1) `decombine` で `src/` へ展開 → 2) エディタで修正 → 3) `combine` で `bin/` に反映。

## Articles

Coming soon! Please check back.

## License

This software is released under the MIT License, see [LICENSE.txt](./LICENSE.txt).
