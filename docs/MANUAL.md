# DocumentBase ユーザーマニュアル

DocumentBase v1.2.0

---

## 1. 概要

DocumentBase は、時間に依存しないドキュメント・ナレッジ資産を Excel 上で構造化管理し、Obsidian Vault への Markdown 出力を行うツールです。

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
| **Output All** | 全 DOC- シートを Obsidian に一括出力 |

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
| collection_output_path | path | 出力先の個別指定（空欄時は自動導出） |

### 5.2 Tbl:DOC_Action

シート上の操作ボタンです。

| ボタン | 動作 |
|--------|------|
| Obsidian へ出力する | 当該 Collection を Obsidian に出力 |
| Refresh (導出値を更新) | no, document_id 等を再計算 |
| Add Collection Sheet | 新しい Collection を追加 |

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

## 7. Obsidian への出力

### 単一 Collection の出力

Collection シート上で **Obsidian へ出力する** ボタンを押します。

### 一括出力

UI_Dashboard の **Output All** ボタンで全 DOC- シートを一括出力します。

### 出力先の決定

以下の優先順位で出力先ディレクトリが決まります:

1. **collection_output_path が設定されている場合** → そのパスを使用
2. **未設定の場合** → `DEF_Parameter.OUTPUT_ROOT / <collection_id>_<collection_name>`

```
例:
OUTPUT_ROOT = D:\ObsidianVault\40_DocumentBase

出力先: D:\ObsidianVault\40_DocumentBase\DOC-TECH-01_Git運用ガイド\
```

出力先フォルダが存在しない場合は自動作成されます。

### 出力ファイル

| 対象 | ファイル名 | 内容 |
|------|-----------|------|
| Collection | `README.md` | collection_id, collection_name, collection_domain, collection_status 等 |
| 各ドキュメント | `<document_id>_<version>_<title>.md` | DOC_DocumentList の行プロパティ |

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

これにより、Obsidian 側で加筆した内容が Excel からの再出力で消えることはありません。

### 出力モード

DEF_Parameter の `OUTPUT_MODE` で制御します。

| モード | 動作 |
|--------|------|
| `auto` | 上記のパス解決ルールに従い自動出力（デフォルト） |
| `dialog` | フォルダ選択ダイアログで任意の場所に出力 |

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
| OUTPUT_ROOT | path | *(空)* | Obsidian 出力ルートディレクトリ |
| OUTPUT_MODE | select | auto | 出力モード（auto / dialog） |
| VERSION_DEFAULT | string | v1.1 | 新規ドキュメントのデフォルトバージョン |
| COLLECTION_TEMPLATE | string | TPL_DOC-CATEGORY-SEQ | テンプレートシート名 |
| PREFIX_COLLECTION | string | DOC- | Collection シート接頭辞 |
| DATA_EXPORT_PATH | path | *(空)* | YAML データエクスポート先パス |

**初回セットアップ時に最低限 `OUTPUT_ROOT` を設定してください。**

---

## 12. 典型的なワークフロー

### 初回セットアップ

1. DEF_Parameter の `OUTPUT_ROOT` に Obsidian Vault のパスを設定
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
4. Obsidian へ出力 → Markdown ファイルが生成される
       ↓
5. Obsidian 側でドキュメントの本文を執筆
       ↓
6. 再出力 → frontmatter だけ更新、本文は保持
```

### 定期メンテナンス

- **Update Index**: Dashboard の統計情報と Collection 一覧を最新化
- **Refresh All**: 全 Collection の導出値を一括再計算
- **Output All**: 全 Collection を Obsidian に一括出力

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
| Pst_OutputToObsidian | Obsidian Markdown 出力（単一・一括） |
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
| パス管理 | 個別登録制 | 自動導出 + 個別上書き可 |
| 出力先導出 | 個別登録必要 | collection_id + name から自動 |

---

## 16. トラブルシューティング

### Refresh しても document_id が生成されない

- **title が入力されているか確認**: title が空の行は無視されます
- **doc_type が入力されているか確認**: doc_type が空の場合、prefix と document_id は空になります
- **DEF_DocType に該当する doc_type が登録されているか確認**

### 出力先フォルダが作成されない

- DEF_Parameter の `OUTPUT_ROOT` が設定されているか確認
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
