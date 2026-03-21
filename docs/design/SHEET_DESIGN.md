# DocumentBase シート設計

## 1. ブック構成

```
DocumentBase.xlsm
├── DEF_PropertyMaster       ← M_* ヘッダーから自動生成
├── DEF_RolePropertyMatrix   ← role × property 適用マトリクス
├── M_Role                   ← プロパティマスタ群
├── M_ProjectGroup           │
├── M_Project                │  （IMG DocumentProperty.xlsx と同一構造）
├── M_ProjectStatus          │
├── M_DoctrineLevel          │
├── M_DocType                │
├── M_DocStatus              │
├── M_Phase                  ←┘
├── IDX_DocumentBase         ← 全 Series の一覧（Index）
└── <Series シート群>         ← MAN-DEV-01, STD-NW-01, ...
```

---

## 2. Series シート設計

### 2.1 シート命名規則

```
<DOC_TYPE>-<CATEGORY>-<No>
```

| 要素 | 由来 | 例 |
|------|------|----|
| DOC_TYPE | M_DocType の value から接頭辞を導出 | MAN, STD, REF, FAQ, POL, CHK |
| CATEGORY | ドキュメントドメイン略称 | DEV, NW, OBS, AI, OPS |
| No | Series 通番（01〜） | 01, 02 |

例:
- `MAN-DEV-01` — 開発系マニュアル #01
- `STD-OPS-01` — 運用標準 #01
- `REF-OBS-01` — Obsidian リファレンス #01

### 2.2 シート構造

FlowBase のプロジェクトシートと同様、ヘッダー情報テーブル + ドキュメントテーブルで構成する。

```
┌─────────────────────────────────────────────────────┐
│  MAN-DEV-01                                          │
│                                                      │
│  Tbl:HeaderInfo                                      │
│  ┌──────────────┬──────────────────────────────────┐ │
│  │ series_id    │ MAN-DEV-01                       │ │
│  │ series_name  │ 開発マニュアル                     │ │
│  │ doc_type     │ manual                           │ │
│  │ category     │ DEV                              │ │
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
| 16 | document_id | string | 導出値: Series ID + no → `MAN-DEV-01-S01` |

### 3.3 document_id の導出規則

```
document_id = <series_id> + "-S" + <no（ゼロ埋め2桁）>
```

例:
- Series `MAN-DEV-01`、no=1 → `MAN-DEV-01-S01`
- Series `STD-NW-01`、no=3 → `STD-NW-01-S03`

### 3.4 file_name の生成規則（出力時のみ）

```
file_name = <document_id> + "_" + <version> + "_" + <title> + ".md"
```

例: `MAN-DEV-01-S01_v1.1_Git運用手引書.md`

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

FlowBase 準拠。テーブル開始位置を示すマーカーとして、テーブル直上のセルに記載する。

```
Tbl:HeaderInfo     ← マーカー（テーブル直上のセル）
┌──────────┬───────┐
│ key      │ value │  ← テーブルヘッダー行
├──────────┼───────┤
│ ...      │ ...   │
```

VBA はマーカー文字列 `Tbl:` を検索してテーブルの開始位置を特定する。

---

## 6. FlowBase との構造対比

| 概念 | FlowBase | DocumentBase |
|------|----------|-------------|
| シート単位 | Project | Series |
| 行単位 | Task | Document |
| 子要素 | — | Section（Obsidian 側） |
| シート名 | PJ-Technology-25-01 | MAN-DEV-01 |
| 行ID | T001 | S01 |
| テーブルマーカー | Tbl: | Tbl:（共通） |
| ヘッダーテーブル | Tbl:HeaderInfo | Tbl:HeaderInfo（共通） |
| メインテーブル | Tbl:TaskList | Tbl:DocumentList |
