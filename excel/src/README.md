# DocumentBase VBA モジュール構成

## ディレクトリ構造

```
src/vba/
├── common/                    ← 共通モジュール（FlowBase 準拠）
│   ├── Mod_Constants.bas      ← 定数定義（Tblマーカー名、シート名、接頭辞）
│   ├── Utl_Logger.bas         ← ログ出力（Debug.Print）
│   ├── Utl_Table.bas          ← テーブル操作（Tblマーカー検索、読み書き）
│   ├── Utl_Sheet.bas          ← シート操作（フィルタ、存在確認、コピー、ソート）
│   ├── Utl_File.bas           ← ファイル操作（FSO、UTF-8読み書き）
│   └── Utl_Yaml.bas           ← YAML シリアライズ／パース
└── tools/                     ← 機能モジュール
    ├── Pst_UpdateIndex.bas           ← UpdateIndex（Dashboard + CollectionIndex 更新）
    ├── Pst_AddCollectionSheet.bas    ← AddCollectionSheet（Collection 追加）
    ├── Pst_RefreshDocumentList.bas   ← RefreshDocumentList（導出値再計算）
    ├── Pst_OutputToObsidian.bas      ← OutputToObsidian（Obsidian .md 出力）
    └── Pst_DataIO.bas                ← DataIO（YAML エクスポート／インポート）
```

## インポート順序

1. `Mod_Constants.bas` — 他モジュールが参照する定数
2. `Utl_Logger.bas` — ログ関数
3. `Utl_Table.bas` — テーブル操作（Mod_Constants に依存）
4. `Utl_Sheet.bas` — シート操作（Mod_Constants, Utl_Table に依存）
5. `Utl_File.bas` — ファイル操作（独立）
6. `Pst_UpdateIndex.bas` — ツール（全共通モジュールに依存）
7. `Pst_AddCollectionSheet.bas` — ツール（全共通モジュールに依存）
8. `Pst_RefreshDocumentList.bas` — ツール（Mod_Constants, Utl_Table, Utl_Sheet に依存）
9. `Pst_OutputToObsidian.bas` — ツール（全共通モジュール + Utl_File に依存）

## UpdateIndex の動作

`Pst_UpdateIndex.UpdateIndex` は以下の3つを一括実行する:

1. **UI_Dashboard `Tbl:UI_SheetIndex`** — 全シート一覧を収集・書き込み
2. **UI_Dashboard `Tbl:UI_Status`** — 統計情報を更新（total_collections, total_documents, active_collections, last_updated）
3. **UI_CollectionIndex `Tbl:CollectionIndex`** — DOC- シートの HeaderInfo から Collection 一覧を収集・書き込み

## AddCollectionSheet の動作

`Pst_AddCollectionSheet.AddCollectionSheet` の処理フロー:

1. `UI_AddSheet` の `Tbl:AddCollection` から入力値を読み取り
2. `collection_name`（必須）、`domain`（必須）をバリデーション
3. domain から短縮コードを導出（Technology→TECH, Mind→MIND 等）
4. 既存 DOC-シートを走査して最大 SEQ を取得 → +1
5. `collection_id` = `DOC-<CODE>-<SEQ>` を生成
6. `TPL_DOC-CATEGORY-SEQ` をコピーしてリネーム
7. `Tbl:DOC_HeaderInfo` に値をセット:
   - **自動**: collection_id, status=active, created=today, updated=today
   - **入力**: collection_name, domain, related_project, summary
8. `Tbl:AddCollection` の入力値をクリア

## RefreshDocumentList の動作

`Pst_RefreshDocumentList.RefreshDocumentList` — アクティブな DOC- シートで実行:

1. `Tbl:DOC_HeaderInfo` から `collection_id` を取得
2. `DEF_DocType` の `Tbl:DEF_DocTypeData` から `doc_type` → `id_prefix` マッピングを読み込み
3. `Tbl:DOC_DocumentList` の各行を走査:
   - `doc_type` → `doc_type_prefix` 列を更新
   - 同一 `doc_type` の出現順をカウントし `document_id` を再生成
     （例: `DOC-TECH-01-M01`, `DOC-TECH-01-RF01`）
4. `Tbl:DOC_HeaderInfo` の `updated` を今日の日付に更新

## OutputToObsidian の動作

`Pst_OutputToObsidian.OutputToObsidian` — アクティブな DOC- シートで実行:

1. `Tbl:DOC_HeaderInfo` を読み取り
2. 出力先パスを解決:
   - `output_path` が非空 → そのまま使用
   - 空 → `DEF_Parameter.OUTPUT_ROOT` / `collection_id`
3. Collection README を出力:
   - ファイル: `<output_dir>/README.md`
   - YAML frontmatter: collection_id, collection_name, domain, related_project, status, created, updated
   - 既存ファイルがあれば frontmatter のみ更新（本文は保持）
4. 各ドキュメント行を出力:
   - ファイル: `<output_dir>/<document_id>_<version>_<title>.md`
   - YAML frontmatter: DOC_DocumentList の全カラム（no を除く）
   - 既存ファイルがあれば frontmatter のみ更新（本文は保持）
5. `Tbl:DOC_HeaderInfo` の `updated` を今日の日付に更新
