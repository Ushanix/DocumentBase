# DocumentBase ユーザーマニュアル

DocumentBase v1.3.0

---

## 1. 概要

DocumentBase は、時間に依存しないドキュメント・ナレッジ資産を Excel 上で構造化管理し、Markdown 出力を行うツールです。出力先は Obsidian Vault に限らず、任意のフォルダを指定できます。

Base シリーズの一つとして、FlowBase（ワークフロー管理）・TimeBase（時間管理）と共通の設計思想を持ちます。

### 基本概念

| 用語 | 説明 | FlowBase 対応 |
|------|------|---------------|
| **Collection** | 同一トピックのドキュメント群。1シート = 1 Collection。書籍に相当。 | Project |
| **Document** | Collection 内の個別ドキュメント。テーブルの1行。章に相当。 | Task |

---

## 2. ブック構成

DocumentBase は以下のシート群で構成されます。

| 接頭辞 | 役割 | 例 |
|--------|------|----|
| `UI_` | 操作画面・フォーム | UI_Dashboard, UI_AddSheet |
| `DEF_` | プロパティ定義・設定 | DEF_DocType, DEF_Parameter |
| `TPL_` | テンプレート | TPL_DOC-CATEGORY-SEQ |
| `DOC-` | Collection シート（実データ） | DOC-TECH-01, DOC-ENV-01 |
| `LOG_` | 更新履歴 | LOG_UpdateHistory |

---

## 3. 操作画面

### 3.1 UI_Dashboard

メイン操作画面です。以下のボタンと情報を持ちます。

#### 操作ボタン

| ボタン | 動作 |
|--------|------|
| **Add Collection Sheet** | UI_AddSheet を開く |
| **Update Index** | シート一覧・統計・Collection 一覧を一括更新 |
| **Refresh All** | 全 DOC- シートの導出値（no, document_id 等）を一括再計算 |
| **Output All** | 全 DOC- シートをノート出力（一括） |

#### 統計情報（Tbl:UI_Status）

- total_collections: Collection の総数
- total_documents: ドキュメントの総数
- active_collections: ステータスが active の Collection 数
- last_updated: 最終更新日時

#### シート索引（Tbl:UI_SheetIndex）

ブック内の全シートが一覧表示されます。Update Index で自動更新されます。

### 3.2 UI_CollectionIndex

Collection の横断一覧です。各 Collection の名称・ドメイン・ドキュメント数・ステータスを確認できます。

### 3.3 UI_AddSheet

新しい Collection を追加するためのフォームです。詳細は「4. Collection の作成」を参照してください。

### 3.4 UI_DataIO

YAML 形式でのデータエクスポート・インポートを行う画面です。詳細は「8. データ入出力」を参照してください。

---

## 4. Collection の作成

### 手順

1. **UI_AddSheet** を開く（Dashboard の「Add Collection Sheet」ボタン、または直接シートを選択）
2. フォーム（Tbl:AddCollection）に以下を入力:

| フィールド | 必須 | 説明 |
|------------|------|------|
| collection_name | Yes | Collection の名称（最大40文字） |
| domain | Yes | 主題領域（ドロップダウンから選択） |
| related_project | No | FlowBase プロジェクトへの参照 |
| summary | No | Collection の概要 |

3. **Add Collection Sheet** ボタンを押す

### 自動処理

ボタン押下時に以下が自動実行されます:

1. collection_name のバリデーション（文字数・禁止文字のチェック）
2. domain から domain_code を導出（例: Technology → TECH）
3. 同一ドメイン内の最大連番を検索し、次の番号を決定
4. collection_id を生成（例: `DOC-TECH-01`）
5. テンプレートをコピーして新シートを作成
6. DOC_HeaderInfo に以下を自動設定:
   - collection_id, collection_name, collection_domain, collection_summary
   - collection_status = `active`
   - collection_created = 今日の日付
   - collection_updated = 今日の日付

### collection_name の制約

| ルール | 制限 |
|--------|------|
| 最大文字数 | 40文字 |
| 禁止文字 | `\ / : * ? " < > \|` |
| 制御文字 | 改行・タブ等は不可 |
| 連続スペース | 不可 |
| 末尾ピリオド | 不可（Windows フォルダ制限） |

### ドメインと domain_code

domain_code は DEF_CollectionDomain シートの `domain_code` 列で管理されます。

| domain | domain_code | シート名例 |
|--------|-------------|------------|
| Technology | TECH | DOC-TECH-01 |
| Env | ENV | DOC-ENV-01 |
| Mind | MIND | DOC-MIND-01 |

新しいドメインを追加するには、DEF_CollectionDomain シートの Tbl:DEF_CollectionDomainData にレコードを追加してください。

---

## 5. Collection シートの構造

各 DOC-* シートは3つのテーブルで構成されます。

### 5.1 Tbl:DOC_HeaderInfo

Collection の書誌情報です。キーと値のペアで構成されます。

| キー | 型 | 説明 |
|------|------|------|
| collection_id | string | Collection ID（シート名と同一、自動設定） |
| collection_name | string | Collection の名称 |
| collection_summary | string | 概要・目的 |
| collection_domain | select | 主題領域 |
| collection_related_project | string | FlowBase プロジェクト参照（任意） |
| collection_status | select | active / done / archived |
| collection_created | date | 作成日（自動設定） |
| collection_updated | date | 最終更新日（Refresh / Output で自動更新） |
| folder_output_path | path | 出力先フォルダの絶対パス（Priority 1、SelectOutputFolder で設定） |
| obsidian_path_form_vault_folder | path | Obsidian Vault からの相対パス（Priority 2） |
| use_folder_output | select | YES / NO — ドキュメントごとにサブフォルダを作成するか |
| collection_output_path | path | *(レガシー)* folder_output_path 未設定時のフォールバック |

### 5.2 Tbl:DOC_Action

シート上の操作ボタンです。

| action_name | ラベル | 対応 Sub | 備考 |
|-------------|--------|----------|------|
| output_notes | フォルダ・ノートを出力する | `OutputNotes` | |
| refresh | Refresh (導出値を更新) | `RefreshDocumentList` | |
| select_output_folder | 出力先フォルダを選択 | `SelectOutputFolder` | FlowBase 共通 |
| open_output_folder | 出力先フォルダを開く | `OpenOutputFolder` | FlowBase 共通 |
| add_collection_sheet | Add Collection Sheet | `AddCollectionSheet` | DOC- シートのみ（TPL 除外） |

### 5.3 Tbl:DOC_DocumentList

ドキュメントの一覧テーブルです。1行 = 1ドキュメント。

| # | カラム | 型 | 入力 | 説明 |
|---|--------|------|------|------|
| 1 | no | number | 自動 | Collection 内通番 |
| 2 | title | string | 手動 | ドキュメントタイトル（**行の存在判定に使用**） |
| 3 | doc_type | select | 手動 | ドキュメント種別（ドロップダウン） |
| 4 | doc_type_prefix | string | 自動 | doc_type の ID 接頭辞 |
| 5 | status | select | 手動 | draft / review / active / done / archived |
| 6 | role | select | 自動 | 空欄なら `docs` が自動設定 |
| 7 | owner_primary | select | 手動 | ドキュメント担当者 |
| 8 | version | string | 手動 | バージョン（例: v1.0） |
| 9 | created | date | 自動 | 空欄なら今日の日付が自動設定 |
| 10 | updated | date | 自動 | 空欄なら今日の日付が自動設定 |
| 11 | tags | string | 手動 | タグ（カンマ区切り） |
| 12 | summary | string | 手動 | 概要 |
| 13 | document_id | string | 自動 | 導出値: `DOC-TECH-01-M01` 等 |

#### ドキュメントの追加方法

1. **title 列** にドキュメントのタイトルを入力
2. **doc_type** をドロップダウンから選択
3. 必要に応じて status, version, tags, summary を入力
4. **Refresh** ボタンを押す → no, doc_type_prefix, document_id が自動生成される

title が入力された行のみが有効行として認識されます。title が空の行は Refresh 時にクリアされます。

#### document_id の導出規則

```
document_id = <collection_id> + "-" + <doc_type_prefix> + <種別内通番（ゼロ埋め2桁）>
```

通番は同一 doc_type ごとにカウントされます。

| no | title | doc_type | document_id |
|----|-------|----------|-------------|
| 1 | Git運用手引書 | manual | DOC-TECH-01-M01 |
| 2 | Gitコマンドリファレンス | reference | DOC-TECH-01-RF01 |
| 3 | ブランチ運用チェックリスト | checklist | DOC-TECH-01-C01 |
| 4 | Git FAQ | manual | DOC-TECH-01-M02 |

---

## 6. Refresh（導出値の更新）

### 単一シートの Refresh

Collection シート上で **Refresh (導出値を更新)** ボタンを押すと、以下が自動計算されます:

| カラム | 処理 |
|--------|------|
| no | 有効行に1から連番を付与 |
| doc_type_prefix | DEF_DocType の id_prefix 列から取得 |
| document_id | collection_id + prefix + 種別内通番 |
| role | 空欄なら `docs` を設定 |
| created | 空欄なら今日の日付を設定 |
| updated | 空欄なら今日の日付を設定 |

collection_updated も自動で今日の日付に更新されます。

### 一括 Refresh

UI_Dashboard の **Refresh All** ボタンで全 DOC- シートを一括処理します。

---

## 7. ノート出力

v1.3.0 で `Pst_OutputToObsidian` は `Pst_OutputNotes` にリネームされました。出力先は Obsidian に限定されず、3階層のパス解決で柔軟に設定できます。

### 単一 Collection の出力

Collection シート上で **ノート出力** ボタンを押します。

### 一括出力

UI_Dashboard の **Output All** ボタンで全 DOC- シートを一括出力します。

### 出力先の決定（3階層パス解決）

以下の優先順位で出力先ディレクトリが決まります:

| 優先度 | ソース | 説明 |
|--------|--------|------|
| **Priority 1** | `folder_output_path` | HeaderInfo に設定された絶対パス。SelectOutputFolder で設定可能。そのまま使用。 |
| **Priority 2** | `OBSIDIAN_PATH_FROM_SYSTEM_ROOT` + `obsidian_path_form_vault_folder` | 両方が設定されている場合のみ有効。Obsidian Vault 専用パス。 |
| **Priority 2.5** | `collection_output_path` | レガシー互換。v1.2.0 以前の個別パス指定。 |
| **Priority 3** | `DOCUMENTBASE_OUTPUT_PATH` or `OUTPUT_ROOT` / collection_folder | ベース出力パス + Collection サブフォルダの自動生成。未設定時は ThisWorkbook.Path。 |

```
例（Priority 1）:
  folder_output_path = D:\MyDocs\DOC-TECH-01_Git運用ガイド
  → D:\MyDocs\DOC-TECH-01_Git運用ガイド\

例（Priority 2）:
  OBSIDIAN_PATH_FROM_SYSTEM_ROOT = D:\ObsidianVault
  obsidian_path_form_vault_folder = 40_DocumentBase/DOC-TECH-01
  → D:\ObsidianVault\40_DocumentBase\DOC-TECH-01\

例（Priority 3）:
  DOCUMENTBASE_OUTPUT_PATH = D:\ObsidianVault\40_DocumentBase
  → D:\ObsidianVault\40_DocumentBase\DOC-TECH-01_Git運用ガイド\
```

出力先フォルダが存在しない場合は自動作成されます。

### フォルダ選択・Explorer 連携

| ボタン | 動作 |
|--------|------|
| **Select Output Folder** | フォルダ選択ダイアログで出力先を選択し、`folder_output_path` に保存。Collection 名サブフォルダを自動付与。 |
| **Open Output Folder** | 3階層パス解決と同じロジックで出力先を特定し、Explorer で開く。フォルダ未作成時は作成を確認。 |

### 出力ファイル

#### フラットモード（デフォルト）

```
<output_dir>/
  README.md                           <- Collection ヘッダー
  <document_id>_<version>_<title>.md  <- 各ドキュメント
```

#### フォルダモード（use_folder_output = YES）

```
<output_dir>/
  README.md                                 <- Collection ヘッダー
  <document_id>_<title>/
    <document_id>_<title>.md                <- 各ドキュメントがサブフォルダに配置
```

フォルダモードは HeaderInfo の `use_folder_output` を `YES` に設定すると有効になります。ドキュメントごとに添付ファイルや関連資料を同梱する場合に便利です。

### 出力の分離原則

- Collection の書誌情報は **README.md のみ** に出力
- 個別ドキュメントには **ドキュメント行のプロパティのみ** を出力
- 書籍の各章に書名を繰り返さないのと同じ設計

### YAML frontmatter の例

**README.md:**
```yaml
---
collection_id: DOC-TECH-01
collection_name: Git運用ガイド
collection_domain: Technology
collection_status: active
collection_created: 2026-03-22
collection_updated: 2026-03-24
---

# Git運用ガイド

Gitの基本操作からブランチ戦略・CI連携までをまとめたナレッジ集
```

**DOC-TECH-01-M01_v1.0_Git運用手引書.md:**
```yaml
---
title: Git運用手引書
doc_type: manual
status: active
role: docs
owner_primary: Ushas
version: v1.0
created: 2026-03-22
updated: 2026-03-24
document_id: DOC-TECH-01-M01
---

# Git運用手引書
```

### 既存ファイルの扱い

出力先に同名ファイルが既に存在する場合:
- **frontmatter のみ更新**（Excel 側の最新値で上書き）
- **本文（body）は保持**される

これにより、出力先で加筆した内容が Excel からの再出力で消えることはありません。

---

## 8. データ入出力（YAML）

UI_DataIO から YAML 形式でシートデータのエクスポート・インポートが可能です。

### エクスポート

1. UI_DataIO を開く
2. **Refresh List** でシート一覧を取得
3. Tbl:ExportList で対象シートの `select` 列を `YES` に設定
4. **Export Selected** ボタンを押す

YAML ファイルは DEF_Parameter の `DATA_EXPORT_PATH` に出力されます。

### インポート

1. UI_DataIO を開く
2. **Refresh List** で YAML ファイル一覧を取得
3. Tbl:ImportList で対象ファイルの `select` 列を `YES` に設定
4. `action` 列で create / overwrite / skip を選択
5. **Import Selected** ボタンを押す

---

## 9. doc_type 一覧

DEF_DocType で管理されるドキュメント種別と ID 接頭辞の対応表です。

| doc_type | prefix | 用途 |
|----------|--------|------|
| manual | M | 手引書 |
| procedure | P | 手順書 |
| checklist | C | チェックリスト |
| requirements | R | 要件定義 |
| design | D | 設計書 |
| template | T | テンプレート |
| knowledge | K | ナレッジ |
| reference | RF | リファレンス |
| master | MS | マスタ |
| readme | RM | README |
| index | IX | 索引 |
| dashboard | DB | ダッシュボード |
| meeting_minutes | MM | 議事録 |
| work_log | WL | 作業ログ |
| output_log | OL | 出力記録 |
| reading_notes | RN | 読書ノート |
| q_and_a | QA | Q&A |
| idea | ID | アイデア |
| concept | CN | コンセプト |
| draft | DR | 草稿 |
| plan | PL | 計画 |
| report | RP | レポート |
| daily | DY | 日報 |
| course_of_action | CA | 行動方針 |
| doctrine | DC | ドクトリン |
| charter | CH | 憲章 |
| policy | PO | ポリシー |
| adr | AD | ADR（意思決定記録） |
| essay | ES | エッセイ |
| analysis | AN | 分析 |
| proposal | PR | 提案 |
| review | RV | レビュー |
| deliverable | DL | 成果物 |
| standards | ST | 標準 |

新しい doc_type を追加するには、DEF_DocType シートの Tbl:DEF_DocTypeData にレコードを追加してください。VBA コードの変更は不要です。

---

## 10. ドキュメントステータス

DOC_DocumentList の status 列で管理します。

| status | 表示名 | 説明 |
|--------|--------|------|
| draft | 草稿 | 執筆中・未完成 |
| review | レビュー | レビュー待ち・確認中 |
| active | 有効 | 運用中・参照可能 |
| done | 完了 | 完了・改訂予定なし |
| archived | アーカイブ | アーカイブ済・参照のみ |

Collection のステータス（collection_status）は active / done / archived の3値です。

---

## 11. DEF_Parameter 設定一覧

| パラメータ | 型 | 既定値 | 説明 |
|------------|------|--------|------|
| OUTPUT_ROOT | path | *(空)* | ノート出力ルートディレクトリ（Priority 3 フォールバック） |
| DOCUMENTBASE_OUTPUT_PATH | path | *(空)* | ノート出力ベースパス（Priority 3、OUTPUT_ROOT より優先） |
| OBSIDIAN_PATH_FROM_SYSTEM_ROOT | path | *(空)* | Obsidian Vault のシステムルートパス（Priority 2） |
| OUTPUT_MODE | select | auto | 出力モード（auto / dialog） |
| VERSION_DEFAULT | string | v1.1 | 新規ドキュメントのデフォルトバージョン |
| COLLECTION_TEMPLATE | string | TPL_DOC-CATEGORY-SEQ | テンプレートシート名 |
| PREFIX_COLLECTION | string | DOC- | Collection シート接頭辞 |
| DATA_EXPORT_PATH | path | *(空)* | YAML データエクスポート先パス |

**初回セットアップ時に `DOCUMENTBASE_OUTPUT_PATH` または `OUTPUT_ROOT` を設定してください。** Obsidian 連携を使う場合は `OBSIDIAN_PATH_FROM_SYSTEM_ROOT` も設定します。

---

## 12. 典型的なワークフロー

### 初回セットアップ

1. DEF_Parameter の `DOCUMENTBASE_OUTPUT_PATH`（または `OUTPUT_ROOT`）に出力先パスを設定
2. DEF_CollectionDomain にドメインと domain_code を登録（初期値あり）
3. DEF_Owner にユーザー名を登録

### 日常運用

```
1. UI_AddSheet で新 Collection を作成
       ↓
2. DOC-*シートで title と doc_type を入力
       ↓
3. Refresh ボタンで導出値を自動生成
       ↓
4. ノート出力 → Markdown ファイルが生成される
       ↓
5. 出力先でドキュメントの本文を執筆
       ↓
6. 再出力 → frontmatter だけ更新、本文は保持
```

### 定期メンテナンス

- **Update Index**: Dashboard の統計情報と Collection 一覧を最新化
- **Refresh All**: 全 Collection の導出値を一括再計算
- **Output All**: 全 Collection をノート出力

---

## 13. Tbl: マーカー仕様

DocumentBase のすべてのデータ領域は `Tbl:` マーカーで管理されます。

### 規則

- マーカーはテーブル直上のセル（A列）に記載
- VBA は `Tbl:` を検索してテーブルの開始位置を特定
- 形式: `Tbl:<シート接頭辞>_<テーブル名>`

### ケース規則

| 対象 | ケース | 例 |
|------|--------|----|
| シート名 | 接頭辞 + PascalCase | DEF_DocType |
| テーブルマーカー | Tbl: + 接頭辞 + PascalCase | Tbl:DEF_DocTypeData |
| YAML プロパティ | snake_case | doc_type |

### マーカー一覧（主要）

| シート | マーカー | 内容 |
|--------|----------|------|
| DOC-* | Tbl:DOC_HeaderInfo | Collection 書誌情報 |
| DOC-* | Tbl:DOC_Action | 操作ボタン |
| DOC-* | Tbl:DOC_DocumentList | ドキュメント一覧 |
| UI_Dashboard | Tbl:UI_Operations | 操作ボタン |
| UI_Dashboard | Tbl:UI_Status | 統計情報 |
| UI_Dashboard | Tbl:UI_SheetIndex | シート索引 |
| UI_AddSheet | Tbl:AddCollection | 入力フォーム |
| UI_CollectionIndex | Tbl:CollectionIndex | Collection 一覧 |
| DEF_Parameter | Tbl:DEF_Parameter | パラメータ設定 |

---

## 14. VBA モジュール構成

### 共通層（Common）

| モジュール | 役割 |
|------------|------|
| Mod_Constants | 定数定義（マーカー名、シート名、接頭辞、パラメータキー） |
| Utl_Sheet | シート操作（検索、コピー、テーブルリネーム） |
| Utl_Table | テーブル操作（マーカー検索、読み書き、ルックアップ） |
| Utl_File | ファイル操作（UTF-8 読み書き、フォルダ作成、ファイル名サニタイズ） |
| Utl_Logger | ログ出力（イミディエイトウィンドウへ INFO/WARN/ERROR） |
| Utl_Yaml | YAML 変換（シリアライズ・パース） |

### ツール層（Presentation）

| モジュール | 機能 |
|------------|------|
| Pst_AddCollectionSheet | Collection シートの新規作成 |
| Pst_RefreshDocumentList | 導出値の再計算（単一・一括） |
| Pst_OutputNotes | ノート出力（単一・一括）— 3階層パス解決・フォルダモード対応 |
| Pst_FolderActions | フォルダ選択・Explorer 連携（SelectOutputFolder / OpenOutputFolder） |
| Pst_UpdateIndex | Dashboard・CollectionIndex の更新 |
| Pst_DataIO | YAML エクスポート・インポート |

---

## 15. FlowBase との対比

| 概念 | FlowBase | DocumentBase |
|------|----------|-------------|
| シート単位 | Project | Collection |
| 行単位 | Task | Document |
| シート名 | PJ-INFRA-26-01 | DOC-TECH-01 |
| 行ID | T001 | M01, RF01（doc_type別） |
| ヘッダーテーブル | Tbl:PJ_HeaderInfo | Tbl:DOC_HeaderInfo（collection_ 接頭辞） |
| メインテーブル | Tbl:PJ_TaskList | Tbl:DOC_DocumentList |
| パス管理 | 3階層パス解決 | 3階層パス解決（同一設計） |
| 出力先導出 | folder_output_path > Obsidian > base | folder_output_path > Obsidian > base |

---

## 16. トラブルシューティング

### Refresh しても document_id が生成されない

- **title が入力されているか確認**: title が空の行は無視されます
- **doc_type が入力されているか確認**: doc_type が空の場合、prefix と document_id は空になります
- **DEF_DocType に該当する doc_type が登録されているか確認**

### 出力先フォルダが作成されない

- 3階層パス解決のいずれかが設定されているか確認:
  - HeaderInfo の `folder_output_path`（Priority 1）
  - `OBSIDIAN_PATH_FROM_SYSTEM_ROOT` + `obsidian_path_form_vault_folder`（Priority 2）
  - `DOCUMENTBASE_OUTPUT_PATH` または `OUTPUT_ROOT`（Priority 3）
- パスに書き込み権限があるか確認
- ネットワークドライブの場合、接続状態を確認

### ドロップダウンが表示されない

- 対応する DEF_* シートにデータが登録されているか確認
- 名前の定義（lst_*）が正しく設定されているか確認
  - 確認方法: 数式タブ → 名前の管理

### Collection 作成時に「domain_code not resolved」エラー

- DEF_CollectionDomain シートの Tbl:DEF_CollectionDomainData に、入力した domain に対応する `domain_code` 列の値が設定されているか確認

---

## 17. Git 管理

VBA コードは `excel/src/vba/` にエクスポートされ Git で管理されます。Excel ブック（.xlsm）も Git で管理されますが、以下は .gitignore で除外されています:

- `*.py` — Python セットアップスクリプト
- `*.bak` — バックアップファイル
- `export/`, `output/` — 出力ディレクトリ
