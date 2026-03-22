# DocumentBase シート設計

## 1. ブック構成

```
DocumentBase_v0.1.1.xlsm
│
├── UI_AddSheet              ← Collection 追加フォーム + ボタン
├── UI_Dashboard             ← メイン操作画面（統計・シート索引・一括操作）
├── UI_CollectionIndex       ← Collection 横断一覧
├── UI_DataIO                ← YAML エクスポート／インポート
│
├── DEF_CollectionName       ← プロパティ定義群（select / prompt / auto）
├── DEF_Title                │  DocumentProperty.xlsx の M_* から DEF_ に移植
├── DEF_DoctrineLevel        │  select 型はヘッダー + 選択肢データを持つ
├── DEF_CollectionSummary    │  prompt / auto 型はヘッダーのみ
├── DEF_Role                 │
├── DEF_CollectionDomain     │
├── DEF_DocType              │
├── DEF_RelatedProject       │
├── DEF_ProjectGroup         │
├── DEF_CollectionStatus     │
├── DEF_Project              │
├── DEF_CollectionCreated    │
├── DEF_CollectionUpdated    │
├── DEF_Phase                │
├── DEF_Domain               │
├── DEF_Version              │
├── DEF_Created              │
├── DEF_Updated              │
├── DEF_Tags                 │
├── DEF_Summary              │
├── DEF_Owner                ←┘ 管理者（FlowBase owner_primary と共通）
├── DEF_Parameter            ← 出力設定（OUTPUT_ROOT 等）
├── DEF_SheetPrefix          ← シート接頭辞のソート順定義
│
├── TPL_DOC-CATEGORY-SEQ     ← Collection テンプレート
│
├── DOC-TECH-01              ← Collection シート群
├── DOC-TECH-02              │
├── DOC-ENV-01               │
├── DOC-MIND-01              ←┘
│
└── LOG_UpdateHistory        ← 更新履歴
```

### 設計上の判断

- **M_* ではなく DEF_* を採用**: DocumentProperty.xlsx（IMG リポジトリ）が M_* の正本。DocumentBase 内のプロパティシートはブック固有の定義・設定（DEF_）として位置づける。
- **DEF_PropertyMaster / DEF_RolePropertyMatrix は DocumentBase に含めない**: DocumentProperty.xlsx 側で管理し、DEF_ シートの生成元として使用する。
- **IDX_DocumentBase は設けない**: UI_CollectionIndex が Collection 横断一覧を担う。FlowBase の UI_ProjectIndex に相当。

---

## 2. Collection シート設計

### 2.1 シート命名規則

```
DOC-<domain_code>-<No>
```

| 要素 | 由来 | 例 |
|------|------|----|
| DOC | Base 識別子（FlowBase の PJ に相当） | DOC |
| domain_code | DEF_CollectionDomain の value から導出 | TECH, ENV, MIND |
| No | domain 内通番（2桁、拡張時3桁） | 01, 02 |

例:
- `DOC-TECH-01` — Git運用ガイド
- `DOC-TECH-02` — ネットワーク基礎
- `DOC-ENV-01` — PC環境構築
- `DOC-MIND-01` — 読書ノート

### 2.2 シート構造

Collection は書籍に相当する概念であり、ヘッダー情報テーブル（書誌情報）+ アクションテーブル + ドキュメントテーブル（目次・章一覧）で構成する。

```
┌──────────────────────────────────────────────────────────┐
│  DOC-TECH-01                                              │
│                                                           │
│  Tbl:DOC_HeaderInfo                                       │
│  ┌──────────────────┬──────────────────────────────────┐  │
│  │ collection_id    │ DOC-TECH-01                      │  │
│  │ collection_name  │ Git運用ガイド                     │  │
│  │ summary          │ Gitの基本操作から…               │  │
│  │ domain           │ Technology                       │  │
│  │ related_project  │ PJ-Technology-26-01              │  │
│  │ status           │ active                           │  │
│  │ owner_primary    │ Ushas                            │  │
│  │ created          │ 2026-03-22                       │  │
│  │ updated          │ 2026-03-22                       │  │
│  │ output_path      │                                  │  │
│  └──────────────────┴──────────────────────────────────┘  │
│                                                           │
│  Tbl:DOC_Action                                           │
│  ┌────┬──────────────────────┬──────────────────────┐    │
│  │ no │ action_id            │ caption              │    │
│  ├────┼──────────────────────┼──────────────────────┤    │
│  │  1 │ output_to_obsidian   │ Output To Obsidian   │    │
│  │  2 │ refresh_document_list│ Refresh Document List│    │
│  │  3 │ add_collection_sheet │ Add Collection Sheet │    │
│  └────┴──────────────────────┴──────────────────────┘    │
│                                                           │
│  Tbl:DOC_DocumentList                                     │
│  ┌────┬────────────────────┬──────────┬─── ─ ─ ──┐       │
│  │ no │ title              │ doc_type │ ...       │       │
│  ├────┼────────────────────┼──────────┼─── ─ ─ ──┤       │
│  │  1 │ Git運用手引書      │ manual   │ ...       │       │
│  │  2 │ Gitコマンドリファレンス │ reference│ ...   │       │
│  └────┴────────────────────┴──────────┴─── ─ ─ ──┘       │
│                                                           │
└──────────────────────────────────────────────────────────┘
```

### 2.3 Tbl:DOC_HeaderInfo カラム

Collection の書誌情報。書籍の表紙・奥付に相当する。

| key | type | description |
|-----|------|-------------|
| collection_id | string | Collection ID（シート名と同一） |
| collection_name | string | Collection の名称（書籍名に相当） |
| summary | string | Collection の概要・目的 |
| domain | select | 主題領域（DEF_CollectionDomain） |
| related_project | string | FlowBase プロジェクトへの参照（任意） |
| status | select | Collection のステータス（active / done / archived） |
| owner_primary | select | Collection の管理者・筆者（DEF_Owner、FlowBase owner_primary と共通） |
| created | date | 作成日 |
| updated | date | 最終更新日 |
| output_path | path | 出力先パスの個別指定（空欄時は DEF_Parameter の OUTPUT_ROOT から自動導出） |

### 2.4 Tbl:DOC_Action カラム

Collection シート上の操作ボタン定義。

| column | description |
|--------|-------------|
| no | 通番 |
| action_id | VBA プロシージャ対応 ID |
| caption | ボタン表示文字列（Title Case） |
| button | ボタンオブジェクト参照 |

---

## 3. Tbl:DOC_DocumentList カラム設計

### 3.1 設計原則

1. **全カラム保持**: role に関わらず全プロパティ列を保持する
2. **左→右の依存方向**: 参照元カラムが左、導出値カラムが右
3. **title が行の存在判定**: title に入力がある行を有効行とする。no は Refresh で自動採番される導出値
4. **document_id が識別子**: Collection ID + doc_type 接頭辞 + 種別内通番から導出
5. **file_name は非保持**: 出力時に生成

### 3.2 カラム一覧

カラムは依存方向に沿って左→右に配置する。

#### 主キー・基本情報（左側）

| # | column | type | description |
|---|--------|------|-------------|
| 1 | no | number | 導出値。Collection 内通番（Refresh で自動採番） |
| 2 | title | string | ドキュメントタイトル |

#### 分類プロパティ

| # | column | type | description |
|---|--------|------|-------------|
| 3 | doctrine_level | select | ドクトリンレベル |
| 4 | doc_type | select | ドキュメント種別 |
| 5 | doc_type_prefix | string | 導出値: doc_type の ID 接頭辞（Refresh で自動設定） |
| 6 | role | select | ノートの役割 |
| 7 | phase | select | 情報サイクルフェーズ |
| 8 | domain | select | 主題領域 |

#### メタ情報

| # | column | type | description |
|---|--------|------|-------------|
| 9 | version | string | バージョン |
| 10 | created | date | 作成日 |
| 11 | updated | date | 更新日 |
| 12 | tags | string | タグ（カンマ区切り） |
| 13 | summary | string | 概要 |

#### 導出値（右側）

| # | column | type | description |
|---|--------|------|-------------|
| 14 | document_id | string | 導出値: Collection ID + doc_type接頭辞 + 種別内通番 → `DOC-TECH-01-M01` |

### 3.3 document_id の導出規則

```
document_id = <collection_id> + "-" + <doc_type_prefix> + <種別内通番（ゼロ埋め2桁）>
```

#### ナンバリング原則

通番は `no`（テーブル行の主キー）とは独立に、**同一 Collection 内の同一 doc_type に対して一意にカウント** する。

- `no` = テーブル行順序の主キー（全行で連番）
- `document_id` の通番 = 同一 doc_type 内の出現順序

#### doc_type 接頭辞の対応表

接頭辞は DEF_DocType シートの `id_prefix` 列でマスタ管理する。

| doc_type | prefix | 例 |
|----------|--------|----|
| manual | M | DOC-TECH-01-M01 |
| procedure | P | DOC-TECH-01-P01 |
| checklist | C | DOC-TECH-01-C01 |
| requirements | R | DOC-TECH-01-R01 |
| design | D | DOC-TECH-01-D01 |
| template | T | DOC-TECH-01-T01 |
| knowledge | K | DOC-TECH-01-K01 |
| reference | RF | DOC-TECH-01-RF01 |
| master | MS | DOC-TECH-01-MS01 |
| readme | RM | DOC-TECH-01-RM01 |
| index | IX | DOC-TECH-01-IX01 |
| dashboard | DB | DOC-TECH-01-DB01 |
| meeting_minutes | MM | (collection)-MM01 |
| work_log | WL | (collection)-WL01 |
| output_log | OL | (collection)-OL01 |
| reading_notes | RN | (collection)-RN01 |
| q_and_a | QA | (collection)-QA01 |
| idea | ID | (collection)-ID01 |
| concept | CN | (collection)-CN01 |
| draft | DR | (collection)-DR01 |
| plan | PL | (collection)-PL01 |
| report | RP | (collection)-RP01 |
| daily | DY | (collection)-DY01 |
| course_of_action | CA | (collection)-CA01 |
| doctrine | DC | (collection)-DC01 |
| charter | CH | (collection)-CH01 |
| policy | PO | (collection)-PO01 |
| adr | AD | (collection)-AD01 |
| essay | ES | (collection)-ES01 |
| analysis | AN | (collection)-AN01 |
| proposal | PR | (collection)-PR01 |
| review | RV | (collection)-RV01 |
| deliverable | DL | (collection)-DL01 |
| standards | ST | (collection)-ST01 |

#### 導出例

Collection `DOC-TECH-01` 内:

| no | doc_type | 種別内通番 | document_id |
|----|----------|------------|-------------|
| 1 | manual | 1 | DOC-TECH-01-M01 |
| 2 | reference | 1 | DOC-TECH-01-RF01 |
| 3 | checklist | 1 | DOC-TECH-01-C01 |
| 4 | procedure | 1 | DOC-TECH-01-P01 |

### 3.4 file_name の生成規則（出力時のみ）

```
file_name = <document_id> + "_" + <version> + "_" + <title> + ".md"
```

例: `DOC-TECH-01-M01_v1.0_Git運用手引書.md`

---

## 4. UI シート設計

### 4.1 UI_Dashboard

メイン操作画面。統計情報・シート索引・一括操作ボタンを配置。

| テーブル | 内容 |
|----------|------|
| `Tbl:IndexHeader` | sheet_role=dashboard |
| `Tbl:UI_Operations` | 操作ボタン一覧 |
| `Tbl:UI_Status` | 統計情報（total_collections, total_documents, active_collections, last_updated） |
| `Tbl:UI_SheetIndex` | 全シート一覧（UpdateIndex で自動更新） |

#### UI_Operations アクション

| action_id | caption | VBA | 動作 |
|---|---|---|---|
| add_collection_sheet | Add Collection Sheet | — | UI_AddSheet を開く |
| update_index | Update Index | `Pst_UpdateIndex.UpdateIndex` | SheetIndex + Status + CollectionIndex を一括更新 |
| refresh_all | Refresh All | `Pst_RefreshDocumentList.RefreshAll` | 全 DOC- の導出値を一括更新 |
| output_all | Output All | `Pst_OutputToObsidian.OutputAll` | 全 DOC- を Obsidian に一括出力 |

### 4.2 UI_AddSheet

Collection 追加フォーム。

| テーブル | 内容 |
|----------|------|
| `Tbl:IndexHeader` | sheet_role=config |
| `Tbl:UI_Operations` | Add Collection Sheet ボタン |
| `Tbl:AddCollection` | 入力フォーム（collection_name, domain, related_project, summary） |

`AddCollectionSheet` 実行時に VBA が自動セットする値: collection_id, status=active, created=today, updated=today。

### 4.3 UI_CollectionIndex

Collection 横断一覧。FlowBase の UI_ProjectIndex に相当。

| テーブル | 内容 |
|----------|------|
| `Tbl:IndexHeader` | sheet_role=index |
| `Tbl:Actions` | collection_index_update ボタン |
| `Tbl:CollectionIndex` | Collection 一覧（no, collection_id, collection_name, sheet_name, domain, doc_count, status, updated） |

### 4.4 UI_DataIO

YAML エクスポート／インポート。FlowBase の UI_DataIO に相当。

| テーブル | 内容 |
|----------|------|
| `Tbl:IndexHeader` | sheet_role=dashboard |
| `Tbl:UI_Operations` | Refresh List / Export Selected / Import Selected |
| `Tbl:DataIOConfig` | export_format, data_path |
| `Tbl:ExportList` | DOC-/DEF_/TPL_ シート一覧（select, post_action） |
| `Tbl:ImportList` | YAML ファイル一覧（select, action=create/overwrite/skip） |

---

## 5. Tbl: スタートマーカー仕様

### 5.1 基本規則

全てのデータ領域には `Tbl:` マーカーを必須とする。マーカーなしの浮いたデータは禁止。

#### マーカー形式

```
Tbl:<シート接頭辞>_<テーブル名>
```

テーブル名にシート接頭辞を含めることで、ブック内でテーブル名が一意になる。

#### ケース規則

| 対象 | ケース | 例 |
|------|--------|----|
| シート名 | 接頭辞（大文字） + PascalCase | `DEF_DoctrineLevel`, `DEF_Parameter` |
| テーブル名 | `Tbl:` + 接頭辞 + PascalCase | `Tbl:DEF_DoctrineLevelData`, `Tbl:DOC_HeaderInfo` |
| ListObject 名 | マーカーから `Tbl:` を除去 | `DEF_DoctrineLevelData` |
| 名前の定義 | `lst_` + snake_case | `lst_doctrine_level` |
| YAML プロパティ名 | snake_case | `doctrine_level` |

### 5.2 マーカー配置

テーブル直上のセルにマーカー文字列を記載する。VBA は `Tbl:` を検索してテーブルの開始位置を特定する。

```
Tbl:DOC_HeaderInfo       ← マーカー（テーブル直上のセル）
┌──────────┬────────────┐
│ key      │ value      │  ← テーブルヘッダー行
├──────────┼────────────┤
│ ...      │ ...        │
└──────────┴────────────┘
```

1シート内に複数テーブルがある場合、各テーブルの直上にそれぞれマーカーを配置する。

### 5.3 DocumentBase のテーブルマーカー一覧

| シート | マーカー | 内容 |
|--------|----------|------|
| `UI_Dashboard` | `Tbl:IndexHeader` | シート基本情報 |
| `UI_Dashboard` | `Tbl:UI_Operations` | 操作ボタン一覧 |
| `UI_Dashboard` | `Tbl:UI_Status` | 統計情報 |
| `UI_Dashboard` | `Tbl:UI_SheetIndex` | 全シート索引 |
| `UI_AddSheet` | `Tbl:IndexHeader` | シート基本情報 |
| `UI_AddSheet` | `Tbl:UI_Operations` | 操作ボタン |
| `UI_AddSheet` | `Tbl:AddCollection` | Collection 追加フォーム |
| `UI_CollectionIndex` | `Tbl:IndexHeader` | シート基本情報 |
| `UI_CollectionIndex` | `Tbl:Actions` | 操作ボタン |
| `UI_CollectionIndex` | `Tbl:CollectionIndex` | Collection 横断一覧 |
| `UI_DataIO` | `Tbl:IndexHeader` | シート基本情報 |
| `UI_DataIO` | `Tbl:UI_Operations` | 操作ボタン |
| `UI_DataIO` | `Tbl:DataIOConfig` | エクスポート設定 |
| `UI_DataIO` | `Tbl:ExportList` | エクスポート対象一覧 |
| `UI_DataIO` | `Tbl:ImportList` | インポート対象一覧 |
| `DEF_*`（select 型） | `Tbl:DEF_<Name>Header` | プロパティメタデータ |
| `DEF_*`（select 型） | `Tbl:DEF_<Name>Data` | 選択肢データ |
| `DEF_*`（prompt/auto 型） | `Tbl:DEF_<Name>Header` | プロパティメタデータのみ |
| `DEF_Parameter` | `Tbl:DEF_Parameter` | 出力設定 |
| `DEF_SheetPrefix` | `Tbl:DEF_SheetPrefix` | 接頭辞ソート順 |
| `DOC-*`（Collection） | `Tbl:DOC_HeaderInfo` | Collection ヘッダー |
| `DOC-*`（Collection） | `Tbl:DOC_Action` | 操作ボタン |
| `DOC-*`（Collection） | `Tbl:DOC_DocumentList` | ドキュメント一覧 |
| `TPL_DOC-CATEGORY-SEQ` | `Tbl:DOC_HeaderInfo` | 雛形ヘッダー |
| `TPL_DOC-CATEGORY-SEQ` | `Tbl:DOC_Action` | 雛形操作ボタン |
| `TPL_DOC-CATEGORY-SEQ` | `Tbl:DOC_DocumentList` | 雛形ドキュメント一覧 |
| `LOG_UpdateHistory` | `Tbl:IndexHeader` | シート基本情報 |
| `LOG_UpdateHistory` | `Tbl:LOG_UpdateHistory` | 更新履歴 |

---

## 6. FlowBase との構造対比

| 概念 | FlowBase | DocumentBase |
|------|----------|-------------|
| シート単位 | Project | Collection |
| 行単位 | Task | Document |
| 子要素 | — | Section（Obsidian 側） |
| シート名 | PJ-INFRA-26-01 | DOC-TECH-01 |
| 行ID | T001 | M01, RF01, C01（doc_type別） |
| テーブルマーカー | Tbl:PJ\_\* | Tbl:DOC\_\* |
| ヘッダーテーブル | Tbl:PJ\_HeaderInfo | Tbl:DOC\_HeaderInfo |
| メインテーブル | Tbl:PJ\_TaskList | Tbl:DOC\_DocumentList |
| アクションテーブル | Tbl:Actions | Tbl:DOC\_Action |
| パス管理 | DEF_Parameter + 個別パス登録 | DEF_Parameter + output_path 個別上書き |
| 出力先導出 | 個別登録制 | collection_id + collection_name から自動導出 |
| 一括操作 | UpdateAll | RefreshAll / OutputAll |
| データIO | UI_DataIO (YAML) | UI_DataIO (YAML) |

---

## 7. DEF_Parameter 設計

出力先やブック全体の設定を管理するシート。FlowBase の DEF_Parameter に相当する。

### 7.1 パラメータ一覧

| key | type | example | description |
|-----|------|---------|-------------|
| OUTPUT_ROOT | path | `D:\ObsidianVault\40_DocumentBase` | Obsidian 出力ルートディレクトリ |
| OUTPUT_MODE | select | `auto` | 出力モード（auto / dialog） |
| VERSION_DEFAULT | string | `v1.1` | 新規ドキュメントのデフォルトバージョン |
| COLLECTION_TEMPLATE | string | `TPL_DOC-CATEGORY-SEQ` | Collection テンプレートシート名 |
| PREFIX_COLLECTION | string | `DOC-` | Collection シート接頭辞 |
| DATA_EXPORT_PATH | path | | データエクスポート先パス |

### 7.2 Obsidian 出力設計

#### 出力の分離原則

DocumentBase では **Collection ヘッダーと Document の出力先を分離** する。

書籍の各章にいちいち書名・著者名を繰り返さないのと同様に、
Collection の書誌情報は代表ドキュメント（README）にのみ出力し、
個別 Document には Document 行のプロパティのみを出力する。

| 出力対象 | ソース | 出力ファイル | 含むプロパティ |
|---|---|---|---|
| Collection README | Tbl:DOC_HeaderInfo | `<output_dir>/README.md` | collection_id, collection_name, domain, related_project, status, created, updated |
| 個別 Document | Tbl:DOC_DocumentList の各行 | `<output_dir>/<document_id>_<version>_<title>.md` | DOC_DocumentList 全カラム（no を除く） |

### 7.3 出力パス導出規則

Collection の出力先は以下の優先順位で解決する。

#### パス解決ルール

```
if  Tbl:DOC_HeaderInfo の output_path が有効パス
    → output_path をそのまま使用
else
    → DEF_Parameter の OUTPUT_ROOT / <collection_id>_<collection_name>
```

| 条件 | output_path の値 | 解決結果 |
|---|---|---|
| 通常（空欄） | *(empty)* | `D:\ObsidianVault\40_DocumentBase\DOC-TECH-01_Git運用ガイド` |
| 個別指定 | `C:\Users\me\Documents\日報` | `C:\Users\me\Documents\日報` |
| ネットワーク | `\\server\shared\議事録` | `\\server\shared\議事録` |

自動導出時のフォルダ名は `<collection_id>_<collection_name>` 形式とする。
collection_name のファイルシステム不正文字は `_` に置換される。

#### collection_name のバリデーション

AddCollectionSheet 実行時に collection_name を検査し、問題がある場合はポップアップで修正を促す。

| ルール | 制限 |
|---|---|
| 最大文字数 | 40文字 |
| 禁止文字 | `\ / : * ? " < > \|` |
| 制御文字 | 改行・タブ等は不可 |
| 連続スペース | 不可 |
| 末尾ピリオド | 不可（Windows フォルダ制限） |

#### ファイルパス導出

```
<解決済み出力先> / <file_name>
```

#### 導出例

```
output_root:    D:\ObsidianVault\40_DocumentBase

--- output_path 空欄（自動導出） ---
Collection README:
→ D:\ObsidianVault\40_DocumentBase\DOC-TECH-01_Git運用ガイド\README.md

個別 Document:
→ D:\ObsidianVault\40_DocumentBase\DOC-TECH-01_Git運用ガイド\DOC-TECH-01-M01_v1.0_Git運用手引書.md

--- output_path 個別指定 ---
output_path:    C:\Users\me\Documents\日報

Collection README:
→ C:\Users\me\Documents\日報\README.md

個別 Document:
→ C:\Users\me\Documents\日報\DOC-LIFE-01-DY01_v1.0_2026-03-22.md
```

#### フォルダ自動作成

出力先フォルダが存在しない場合、VBA が `MkDir` で自動作成する。
Collection シートの追加だけで、フォルダ構成も自動的に整う。

### 7.4 出力モード

| mode | 動作 |
|------|------|
| `auto` | パス解決ルールに従い自動出力（デフォルト） |
| `dialog` | フォルダ選択ダイアログで任意の場所に出力 |

`dialog` モードは Obsidian 非利用者や、一時的に別の場所へ出力したい場合に使用する。
`dialog` は output_path の設定に関わらず常にダイアログを表示する。

### 7.5 FlowBase 方式との比較

| 観点 | FlowBase | DocumentBase |
|------|----------|-------------|
| DEF_Parameter に持つもの | root + 各PJの相対パス | root のみ |
| 個別パス登録 | 必要（PJシートに path 列） | 任意（output_path で個別上書き可） |
| root 変更時の影響 | DEF_Parameter 1箇所 | DEF_Parameter 1箇所（個別指定分は影響なし） |
| 新Collection 追加時 | パス登録が必要 | 登録不要（自動導出、必要時のみ個別指定） |
| 非Obsidian利用者 | 出力不可または独自パス不可 | dialog モード or output_path 個別指定で対応 |

---

## 8. VBA アクション対応表

全シートの action_id と VBA プロシージャの対応。

### UI_Dashboard

| action_id | VBA | 動作 |
|---|---|---|
| add_collection_sheet | — | UI_AddSheet を開く |
| update_index | `Pst_UpdateIndex.UpdateIndex` | SheetIndex + Status + CollectionIndex を一括更新 |
| refresh_all | `Pst_RefreshDocumentList.RefreshAll` | 全 DOC- の導出値を一括更新 |
| output_all | `Pst_OutputToObsidian.OutputAll` | 全 DOC- を Obsidian に一括出力 |

### UI_AddSheet

| action_id | VBA | 動作 |
|---|---|---|
| add_collection_sheet | `Pst_AddCollectionSheet.AddCollectionSheet` | Collection シート作成 |

### UI_CollectionIndex

| action_id | VBA | 動作 |
|---|---|---|
| collection_index_update | `Pst_UpdateIndex.CollectionIndexUpdate` | CollectionIndex のみ更新 |

### UI_DataIO

| action_id | VBA | 動作 |
|---|---|---|
| refresh_list | `Pst_DataIO.RefreshList` | ExportList + ImportList を更新 |
| export_selected | `Pst_DataIO.ExportSelected` | 選択シートを YAML にエクスポート |
| import_selected | `Pst_DataIO.ImportSelected` | 選択 YAML をシートにインポート |

### DOC-* (Collection シート)

| action_id | VBA | 動作 |
|---|---|---|
| output_to_obsidian | `Pst_OutputToObsidian.OutputToObsidian` | 当該 Collection を Obsidian に出力 |
| refresh_document_list | `Pst_RefreshDocumentList.RefreshDocumentList` | 当該 Collection の導出値を更新 |
| add_collection_sheet | `Pst_AddCollectionSheet.AddCollectionSheet` | 新しい Collection を追加 |
