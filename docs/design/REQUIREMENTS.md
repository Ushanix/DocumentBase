# DocumentBase 要件定義

## 1. 目的

DocumentBase は、時空に依存しないドキュメント・ナレッジ資産を
Excel 上で構造化管理し、Obsidian Vault への YAML frontmatter 出力を行うツールである。

InformationManagementGuidelines（IMG）で定義されたプロパティ規格に準拠し、
FlowBase・TimeBase と共通の設計思想を持つ Base シリーズの一つとして位置づけられる。

---

## 2. 用語定義

| 用語 | 定義 | FlowBase 対応 |
|------|------|---------------|
| **Series** | 同一の目的・分野に属するドキュメント群。1シート = 1 Series。アーカイブズ学（ISO 15489 / ISAD(G)）における標準術語。 | Project |
| **Document** | Series 内の個別ドキュメント。テーブルの1行に対応。 | Task |
| **Section** | Document の子要素（章・節）。当面は Obsidian 側で管理。 | — |
| **M_\* シート** | プロパティの選択肢を定義するマスタシート。ヘッダー領域にメタ情報を持つ。 | — |
| **DEF_\* シート** | M_\* シートから自動生成される定義・集約ビュー。 | — |

### Series の命名規則

Series のシート名は `<DOC_TYPE>-<CATEGORY>-<No>` 形式とする。

| 要素 | 説明 | 例 |
|------|------|----|
| DOC_TYPE | 文書系列の接頭辞（doc_type から導出） | MAN, STD, REF, FAQ |
| CATEGORY | ドキュメントドメイン | DEV, NW, OBS, AI |
| No | Series 内通番（2桁、拡張時3桁） | 01, 02, ... |

例: `MAN-DEV-01`（開発系マニュアル Series #01）

---

## 3. 機能要件

### 3.1 ドキュメント管理

- **FR-001**: Series シート上でドキュメントの登録・編集・削除ができること
- **FR-002**: テーブルは role に関わらず全プロパティ列を保持すること
- **FR-003**: `no` 列をテーブル内の主キーとすること
- **FR-004**: `document_id` は `doc_type` + `category` + `no` から導出し、導出元カラムより右側に配置すること
- **FR-005**: `file_name` はカラムとせず、出力時に `document_id` + `version` + `title` から生成すること

### 3.2 マスタ管理

- **FR-010**: 各プロパティの定義（メタ情報 + 選択肢）は個別の M_\* シートで管理すること
- **FR-011**: M_\* シートはヘッダー領域（行1-9: メタ情報）とデータ領域（行11+: 選択肢）の統一構造を持つこと
- **FR-012**: DEF_PropertyMaster は全 M_\* シートのヘッダー領域を自動集約して生成すること
- **FR-013**: role の定義は「独自プロパティセットを持つ M_\* シートの存在」から導出されること

### 3.3 Obsidian 連携

- **FR-020**: M_\* シートの定義から Obsidian Templater スクリプトを VBA マクロで自動生成できること
- **FR-021**: 出力される YAML frontmatter は role に応じて必要なプロパティのみを含むこと
- **FR-022**: YAML のプロパティ順序は `yaml_order` に従うこと

### 3.4 FlowBase 共通仕様

- **FR-030**: テーブルには `Tbl:<TableName>` のスタートマーカーを付与すること（FlowBase 準拠）
- **FR-031**: VBA モジュールのうち M_\* シート読取・Templater 生成等の共通処理は BaseCommon として切り出し可能な設計とすること

---

## 4. 非機能要件

- **NFR-001**: Excel ブックのシート数は数十枚以内に収まること
- **NFR-002**: VBA コードは excel/src/ にエクスポートし Git 管理すること
- **NFR-003**: M_\* シートの追加のみで新プロパティを導入でき、VBA コードの変更が不要であること

---

## 5. スコープ外

以下は DocumentBase のスコープ外とし、他の Base または Obsidian 側で管理する。

| 項目 | 管理先 |
|------|--------|
| タスク管理（task_status, priority, estimate 等） | FlowBase |
| 期間計画・日報（period, start_date, end_date, week_of 等） | TimeBase |
| Section（文書の章・節）の構造管理 | Obsidian 側 |
| Web ダッシュボード | 当面なし（将来検討） |

---

## 6. ドキュメントプロパティ一覧

IMG の DocumentProperty.xlsx で定義された全プロパティ。
DocumentBase のテーブルは全列を保持し、role に応じて出力をフィルタする。

### 共通プロパティ（applies_to: all）

| yaml_order | property_name | data_type | input_method |
|------------|---------------|-----------|-------------|
| 1 | title | string | prompt |
| 3 | role | select | suggester |
| 5 | projectGroup | select | suggester |
| 6 | project | select | suggester |
| 13 | phase | select | suggester |
| 17 | version | string | auto |
| 18 | created | date | auto |
| 19 | updated | date | auto |
| 20 | tags | list | manual |
| 21 | summary | string | manual |

### role=project 専用

| yaml_order | property_name | data_type | input_method |
|------------|---------------|-----------|-------------|
| 7 | project_status | select | suggester |

### role=docs, dashboard 専用

| yaml_order | property_name | data_type | input_method |
|------------|---------------|-----------|-------------|
| 2 | doctrine_level | select | suggester |
| 4 | doc_type | select | suggester |
| 9 | status | select | suggester |

---

## 7. 開発フェーズ

### Phase 1: IMG でプロパティ定義（完了）

- DocumentProperty.xlsx による M_\* シート定義
- generate_property_master.py による初版生成

### Phase 2: DocumentBase 開発

- Excel ブック作成（M_\* シート + Series シート）
- VBA: DEF_PropertyMaster 自動集約
- VBA: Templater スクリプト自動生成
- VBA: Obsidian YAML 出力

### Phase 3: IMG に逆導入・再定義

- DocumentBase でプロパティ定義自体を管理（セルフホスティング）
- IMG の Master/ を DocumentBase 出力物に置換
