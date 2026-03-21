# DocumentBase シート設計

## 1. ブック構成

```
DocumentBase.xlsm
├── DEF_Parameter            ← 出力設定（output_root 等）
├── DEF_PropertyMaster       ← M_* ヘッダーから自動生成
├── DEF_RolePropertyMatrix   ← role × property 適用マトリクス
├── M_Role                   ← プロパティマスタ群
├── M_ProjectGroup           │
├── M_Project                │  （IMG DocumentProperty.xlsx と同一構造）
├── M_ProjectStatus          │
├── M_DoctrineLevel          │
├── M_DocType                │
├── M_DocStatus              │
├── M_Phase                  │
├── M_Domain                 ←┘
├── IDX_DocumentBase         ← 全 Series の一覧（Index）
└── <Series シート群>         ← DOC-TECH-01, DOC-ENV-01, ...
```

---

## 2. Series シート設計

### 2.1 シート命名規則

```
DOC-<domain_code>-<No>
```

| 要素 | 由来 | 例 |
|------|------|----|
| DOC | Base 識別子（FlowBase の PJ に相当） | DOC |
| domain_code | M_Domain の value（FlowBase category_code と共通） | TECH, ENV, MIND |
| No | domain 内通番（2桁、拡張時3桁） | 01, 02 |

例:
- `DOC-TECH-01` — Git
- `DOC-TECH-02` — ネットワーク
- `DOC-ENV-01` — PC環境構築
- `DOC-MIND-01` — 読書ノート

### 2.2 シート構造

FlowBase のプロジェクトシートと同様、ヘッダー情報テーブル + ドキュメントテーブルで構成する。

```
┌─────────────────────────────────────────────────────┐
│  DOC-TECH-01                                         │
│                                                      │
│  Tbl:HeaderInfo                                      │
│  ┌──────────────┬──────────────────────────────────┐ │
│  │ series_id    │ DOC-TECH-01                      │ │
│  │ series_name  │ Git                              │ │
│  │ domain       │ Technology                       │ │
│  │ category     │ TECH                             │ │
│  │ owner        │                                  │ │
│  │ status       │ active                           │ │
│  │ created      │ 2026-03-21                       │ │
│  │ updated      │ 2026-03-21                       │ │
│  └──────────────┴──────────────────────────────────┘ │
│                                                      │
│  Tbl:DocumentList                                    │
│  ┌────┬────────┬─────────┬──────────┬─── ─ ─ ──┐    │
│  │ no │ title  │ role    │ doctrine │ ...       │    │
│  │    │        │         │ _level   │           │    │
│  ├────┼────────┼─────────┼──────────┼─── ─ ─ ──┤    │
│  │  1 │ Git…   │ docs    │ manual   │ ...       │    │
│  │  2 │ Docker…│ docs    │ manual   │ ...       │    │
│  └────┴────────┴─────────┴──────────┴─── ─ ─ ──┘    │
│                                                      │
└─────────────────────────────────────────────────────┘
```

---

## 3. Tbl:DocumentList カラム設計

### 3.1 設計原則

1. **全カラム保持**: role に関わらず全プロパティ列を保持する
2. **左→右の依存方向**: 参照元カラムが左、導出値カラムが右
3. **no が主キー**: テーブル内の行識別子。document_id は no から導出
4. **file_name は非保持**: 出力時に生成

### 3.2 カラム一覧

カラムは依存方向に沿って左→右に配置する。

#### 主キー・基本情報（左側）

| # | column | type | description |
|---|--------|------|-------------|
| 1 | no | number | 主キー。Series 内通番。 |
| 2 | title | string | ドキュメントタイトル |
| 3 | role | select | ノートの役割 |

#### 共通プロパティ

| # | column | type | description |
|---|--------|------|-------------|
| 4 | projectGroup | select | プロジェクト大区分 |
| 5 | project | select | プロジェクト |
| 6 | phase | select | 情報サイクルフェーズ |

#### role=project 専用

| # | column | type | description |
|---|--------|------|-------------|
| 7 | project_status | select | プロジェクトステータス（role=project のみ使用） |

#### role=docs, dashboard 専用

| # | column | type | description |
|---|--------|------|-------------|
| 8 | doctrine_level | select | ドクトリンレベル（role=docs/dashboard のみ使用） |
| 9 | doc_type | select | ドキュメント種別（role=docs/dashboard のみ使用） |
| 10 | status | select | ドキュメントステータス（role=docs/dashboard のみ使用） |

#### メタ情報

| # | column | type | description |
|---|--------|------|-------------|
| 11 | version | string | バージョン |
| 12 | created | date | 作成日 |
| 13 | updated | date | 更新日 |
| 14 | tags | string | タグ（カンマ区切り） |
| 15 | summary | string | 概要 |

#### 導出値（右側）

| # | column | type | description |
|---|--------|------|-------------|
| 16 | document_id | string | 導出値: Series ID + doc_type接頭辞 + 種別内通番 → `DOC-TECH-01-M01` |

### 3.3 document_id の導出規則

```
document_id = <series_id> + "-" + <doc_type_prefix> + <種別内通番（ゼロ埋め2桁）>
```

#### ナンバリング原則

通番は `no`（テーブル行の主キー）とは独立に、**同一 Series 内の同一 doc_type に対して一意にカウント** する。

- `no` = テーブル行順序の主キー（全行で連番）
- `document_id` の通番 = 同一 doc_type 内の出現順序

これにより document_id は「同じシリーズの同一ドキュメント種別で何番目か」を表す。
テーブル上で最下位通番が不揃いになるが、行順序は `no` が担保するため問題ない。

#### doc_type 接頭辞の対応表

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
| meeting_minutes | MM | (series)-MM01 |
| work_log | WL | (series)-WL01 |
| output_log | OL | (series)-OL01 |
| reading_notes | RN | (series)-RN01 |
| q_and_a | QA | (series)-QA01 |
| idea | ID | (series)-ID01 |
| concept | CN | (series)-CN01 |
| draft | DR | (series)-DR01 |
| plan | PL | (series)-PL01 |
| report | RP | (series)-RP01 |
| daily | DY | (series)-DY01 |
| course_of_action | CA | (series)-CA01 |
| doctrine | DC | (series)-DC01 |
| charter | CH | (series)-CH01 |
| policy | PO | (series)-PO01 |
| adr | AD | (series)-AD01 |
| essay | ES | (series)-ES01 |
| analysis | AN | (series)-AN01 |
| proposal | PR | (series)-PR01 |
| review | RV | (series)-RV01 |
| deliverable | DL | (series)-DL01 |
| standards | ST | (series)-ST01 |

※ 接頭辞は M_DocType シートに `id_prefix` 列として追加し、マスタ管理する。

#### 導出例

Series `DOC-TECH-01` 内:

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

例: `DOC-TECH-01-M01_v1.1_Git運用手引書.md`

---

## 4. IDX_DocumentBase 設計

全 Series を横断した一覧シート。Series シートの Tbl:HeaderInfo から自動集約する。

| column | description |
|--------|-------------|
| series_id | Series ID |
| series_name | Series 名称 |
| doc_type | 文書系列 |
| category | ドメイン略称 |
| doc_count | 配下ドキュメント数 |
| status | Series ステータス |
| updated | 最終更新日 |

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
| シート名 | 接頭辞（大文字） + PascalCase | `M_ProjectStatus`, `DEF_Parameter` |
| テーブル名 | `Tbl:` + シート接頭辞 + PascalCase | `Tbl:M_PhaseData`, `Tbl:DEF_PropertyMaster` |
| YAML プロパティ名 | snake_case | `project_status`, `doctrine_level` |

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
| `UI_Dashboard` | `Tbl:UI_Operations` | マクロボタン一覧 |
| `UI_Dashboard` | `Tbl:UI_Status` | 統計情報 |
| `UI_Dashboard` | `Tbl:UI_SheetIndex` | M_\* シート索引 |
| `DEF_PropertyMaster` | `Tbl:DEF_PropertyMaster` | プロパティ集約ビュー |
| `DEF_RolePropertyMatrix` | `Tbl:DEF_RolePropertyMatrix` | role × property マトリクス |
| `DEF_Parameter` | `Tbl:DEF_Parameter` | 出力設定 |
| `M_*`（select 型） | `Tbl:M_<Name>Header` | ヘッダー領域 |
| `M_*`（select 型） | `Tbl:M_<Name>Data` | 選択肢データ |
| `M_*`（auto/manual 型） | `Tbl:M_<Name>Header` | ヘッダー領域のみ |
| `DOC-*`（Series） | `Tbl:DOC_HeaderInfo` | Series ヘッダー |
| `DOC-*`（Series） | `Tbl:DOC_DocumentList` | ドキュメント一覧 |
| `IDX_DocumentBase` | `Tbl:IDX_SeriesList` | Series 横断一覧 |
| `TPL_DOC-DOMAIN-NO` | `Tbl:TPL_HeaderInfo` | 雛形ヘッダー |
| `TPL_DOC-DOMAIN-NO` | `Tbl:TPL_DocumentList` | 雛形ドキュメント一覧 |

---

## 6. FlowBase との構造対比

| 概念 | FlowBase | DocumentBase |
|------|----------|-------------|
| シート単位 | Project | Series |
| 行単位 | Task | Document |
| 子要素 | — | Section（Obsidian 側） |
| シート名 | PJ-Technology-25-01 | DOC-TECH-01 |
| 行ID | T001 | M01, RF01, C01（doc_type別） |
| テーブルマーカー | Tbl:PJ\_\* | Tbl:DOC\_\*（接頭辞付き） |
| ヘッダーテーブル | Tbl:PJ\_HeaderInfo | Tbl:DOC\_HeaderInfo |
| メインテーブル | Tbl:PJ\_TaskList | Tbl:DOC\_DocumentList |
| パス管理 | DEF_Parameter + 個別パス登録 | DEF_Parameter（output_root のみ） |
| 出力先導出 | 個別登録制 | Series ID から自動導出 |

---

## 7. DEF_Parameter 設計

出力先やブック全体の設定を管理するシート。FlowBase の DEF_Parameter に相当する。

### 7.1 パラメータ一覧

| key | type | example | description |
|-----|------|---------|-------------|
| output_root | path | `D:\ObsidianVault\40_DocumentBase` | Obsidian 出力ルートディレクトリ |
| output_mode | select | `auto` | 出力モード（auto / dialog） |
| version_default | string | `v1.1` | 新規ドキュメントのデフォルトバージョン |

### 7.2 出力パス導出規則

FlowBase ではプロジェクトごとに出力先パスを個別登録する方式だが、
DocumentBase では Series ID の命名規則が確定的であるため、個別登録を不要とする。

```
output_path = <output_root> / <series_id> / <file_name>
```

#### 導出例

```
output_root:  D:\ObsidianVault\40_DocumentBase
series_id:    DOC-TECH-01
file_name:    DOC-TECH-01-M01_v1.1_Git運用手引書.md

→ D:\ObsidianVault\40_DocumentBase\DOC-TECH-01\DOC-TECH-01-M01_v1.1_Git運用手引書.md
```

#### フォルダ自動作成

出力先フォルダが存在しない場合、VBA が `MkDir` で自動作成する。
Series シートの追加だけで、フォルダ構成も自動的に整う。

### 7.3 出力モード

| mode | 動作 |
|------|------|
| `auto` | output_root + series_id で自動出力（デフォルト） |
| `dialog` | フォルダ選択ダイアログで任意の場所に出力 |

`dialog` モードは Obsidian 非利用者や、一時的に別の場所へ出力したい場合に使用する。

### 7.4 FlowBase 方式との比較

| 観点 | FlowBase | DocumentBase |
|------|----------|-------------|
| DEF_Parameter に持つもの | root + 各PJの相対パス | root のみ |
| 個別パス登録 | 必要（PJシートに path 列） | 不要 |
| root 変更時の影響 | DEF_Parameter 1箇所 | DEF_Parameter 1箇所 |
| 新Series 追加時 | パス登録が必要 | 登録不要（自動導出） |
| 非Obsidian利用者 | 出力不可または独自パス不可 | dialog モードで対応 |
