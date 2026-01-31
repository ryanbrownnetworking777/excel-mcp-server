# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

Excel MCP ServerはModel Context Protocol（MCP）を使用してMicrosoft Excelファイルの読み書きを行うサーバーです。GoとTypeScriptのハイブリッドプロジェクトで、MCPツールとしてExcelデータの操作機能を提供します。

## ビルドコマンド

```bash
# プロジェクト全体のビルド（Go + TypeScript）
npm run build

# TypeScriptのwatch mode
npm run watch

# デバッグ用のMCPインスペクター起動
npm run debug
```

### テスト
```bash
go test ./...                           # 全Goテスト実行
go test ./internal/tools -v             # 特定パッケージのテスト
go test -run TestReadSheetData ./internal/tools  # 特定テスト実行
```

### リンティングとフォーマット
```bash
go fmt ./...      # Goコードのフォーマット
go vet ./...      # Goコードの静的解析
```

## アーキテクチャ

### コアコンポーネント

- **Goバックエンド**: 実際のExcel操作を担当
  - `cmd/excel-mcp-server/main.go`: エントリーポイント
  - `internal/server/server.go`: MCPサーバー実装
  - `internal/excel/`: Excel操作のインターフェースと実装
  - `internal/tools/`: MCPツールの実装

- **TypeScriptランチャー**:
  - `launcher/launcher.ts`: プラットフォーム固有のバイナリを起動
  - Goバイナリの配布とプロセス管理を担当

### Excel操作層

- **Excel Interface**: `internal/excel/excel.go`で定義された抽象インターフェース
- **実装方式**:
  - `excel_excelize.go`: xuri/excelizeライブラリ使用（クロスプラットフォーム）
  - `excel_ole.go`: OLE自動化使用（Windows専用、ライブ編集対応）

### ツール構成

全てのMCPツールは`internal/tools/`に実装：
- `excel_create_file.go`: **新規Excelファイル作成**
- `excel_describe_sheets.go`: シート情報取得
- `excel_read_sheet.go`: データ読み取り（ページネーション対応）
- `excel_write_to_sheet.go`: データ書き込み（新規ファイル対応）
- `excel_format_range.go`: セル書式設定（上流から）
- `excel_create_table.go`: テーブル作成
- `excel_copy_sheet.go`: シートコピー
- `excel_screen_capture.go`: スクリーンキャプチャ（Windows専用）

### 財務モデリング機能（フェーズ1）
- `excel_financial_template.go`: 財務テンプレート
- `excel_circular_reference.go`: 循環参照検出
- `excel_model_validation.go`: モデル検証
- `excel_financial_formatting.go`: 財務書式設定
- `excel_historical_analysis.go`: 履歴分析

### 設定

環境変数での動作制御：
- `EXCEL_MCP_PAGING_CELLS_LIMIT`: ページング時の最大セル数（デフォルト: 4000）

## 新規ファイル作成機能

### 自動ファイル作成
- **全てのツール**が存在しないファイルパスに対して自動的に新規ファイルを作成
- `excel.OpenFile()`が`excel.CreateNewFile()`にフォールバック
- ディレクトリが存在しない場合は自動作成

### 実装詳細
- **Excelizeバックエンド**: `excelize.NewFile()`で新規ワークブック作成
- **OLEバックエンド**: `NewExcelOleWithNewFile()`でExcelアプリケーション経由作成
- **クロスプラットフォーム対応**: macOS/LinuxではExcelizeを使用

## ファイル構造

```
cmd/excel-mcp-server/     # メインアプリケーションエントリーポイント
internal/
  excel/                  # Excel抽象化レイヤー
  server/                 # MCPサーバー実装
  tools/                  # MCPツール実装
launcher/                 # TypeScriptランチャー
memory-bank/              # 開発コンテキストと進捗
```

## ビルドシステム

GoReleaser（`.goreleaser.yaml`）でクロスプラットフォームバイナリを生成：
- Windows: amd64, 386, arm64
- macOS: amd64, arm64
- Linux: amd64, 386, arm64

TypeScriptランチャーは`dist/launcher.js`にコンパイルされ、NPMで公開されます。

## プラットフォーム差異

**Windows専用機能**:
- OLE自動化によるライブExcel操作
- スクリーンキャプチャ機能
- Excelのインストールが必要

**クロスプラットフォーム機能**:
- ファイルベースのExcel操作のみ
- ライブ編集機能なし
- xlsx, xlsm, xltx, xltm形式をサポート

## 依存関係

**Go**: Go 1.23.0以上（Go 1.24.0ツールチェーン推奨）
**Node.js**: TypeScriptコンパイル用にNode.js 20.x以上
**主要パッケージ**:
- `github.com/mark3labs/mcp-go` - MCPフレームワーク
- `github.com/xuri/excelize/v2` - Excelファイル操作
- `github.com/go-ole/go-ole` - Windows OLE自動化

## 開発ワークフロー

- **コミットガイドライン**:
  - 変更セットの完了後、日本語と英語の両言語で説明的なコミットメッセージを作成し、リモートにプッシュする
