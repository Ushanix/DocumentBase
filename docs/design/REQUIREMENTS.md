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
| **Collection** | 同一の主題（トピック）に属するドキュメント群（文書群）。1シート = 1 Collection。doc_type 混成を許容し、トピック単位で束ねる。 | Project |
| **Document** | Collection 内の個別ドキュメント。テーブルの1行に対応。基本数ページの単位。 | Task |
| **Section** | Document の子要素（章・節）。当面は Obsidian 側で管理。 | — |
| **M_\* シート** | プロパティの選択肢を定義するマスタシート。ヘッダー領域にメタ情報を持つ。 | — |
| **DEF_\* シート** | M_\* シートから自動生成される定義・集約ビュー。 | — |

### Collection とプロジェクトの関係

プロジェクト（FlowBase）は Collection を産むが、Collection はプロジェクトから独立して存在する。
書籍の執筆プロジェクトと、成果物としての書籍が別物であるのと同様に、
プロジェクト名と Collection 名は分離される。

- FlowBase: `PJ-Technology-26-01`「TWLM研究プロジェクト」 → 活動の管理
- DocumentBase: `DOC-TECH-03`「TWLMの研究」 → 成果物の管理

Collection ヘッダーには `collection_related_project` として FlowBase プロジェクトへの参照を任意で持てるが、必須ではない。

### Collection の命名規則

Collection のシート名は `DOC-<domain_code>-<No>` 形式とする。

| 要素 | 説明 | 例 |
|------|------|----|
| DOC | Base 識別子（FlowBase の PJ に相当） | DOC |
| domain_code | DEF_CollectionDomain の `domain_code` 列で定義（FlowBase の category_code と共通） | TECH, ENV, MIND |
| No | domain 内通番（2桁、拡張時3桁） | 01, 02 |

例: `DOC-TECH-01`（Technology 領域 Collection #01「Gitの使い方」）

---

## 3. 機能要件

### 3.1 ドキュメント管理

- **FR-001**: Collection シート上でドキュメントの登録・編集・削除ができること
- **FR-002**: テーブルは role に関わらず全プロパティ列を保持すること
- **FR-003**: `no` 列をテーブル内の主キーとすること
- **FR-004**: `document_id` は `collection_id` + `doc_type` の `id_prefix` + 種別内通番から導出し、導出元カラムより右側に配置すること。通番は同一 Collection 内の同一 doc_type に対して一意にカウントする（`no` とは独立）
- **FR-005**: `file_name` はカラムとせず、出力時に `document_id` + `version` + `title` から生成すること

### 3.2 マスタ管理

- **FR-010**: 各プロパティの定義（メタ情報 + 選択肢）は個別の M_\* シートで管理すること
- **FR-011**: M_\* シートはヘッダー領域（行1-9: メタ情報）とデータ領域（行11+: 選択肢）の統一構造を持つこと
- **FR-012**: DEF_PropertyMaster は全 M_\* シートのヘッダー領域を自動集約して生成すること
- **FR-013**: role の定義は「独自プロパティセットを持つ M_\* シートの存在」から導出されること
- **FR-014**: DEF_DocType シートに `id_prefix` 列を持ち、document_id の種別接頭辞をマスタ管理すること
- **FR-015**: DEF_CollectionDomain シートに `domain_code` 列を持ち、シート命名用の短縮コードをマスタ管理すること（ハードコーディング禁止）
- **FR-016**: DOC_DocumentList に `status` 列（draft / review / active / done / archived）を持ち、ドキュメント単位の進捗を管理できること
- **FR-017**: DOC_HeaderInfo の全キーには `collection_` 接頭辞を付与し、DOC_DocumentList のプロパティ名との衝突を防ぐこと

### 3.3 Obsidian 連携

- **FR-020**: M_\* シートの定義から Obsidian Templater スクリプトを VBA マクロで自動生成できること
- **FR-021**: 出力される YAML frontmatter は role に応じて必要なプロパティのみを含むこと
- **FR-022**: YAML のプロパティ順序は `yaml_order` に従うこと
- **FR-023**: 出力先パスは DEF_Parameter の `OUTPUT_ROOT` + Collection ID から自動導出すること（個別パス登録不要、`collection_output_path` で個別上書き可）
- **FR-024**: 出力先フォルダが存在しない場合は自動作成すること
- **FR-025**: 出力モード `dialog` により任意のフォルダへの出力を選択できること（非 Obsidian 利用者対応）
- **FR-026**: Collection ヘッダー（Tbl:DOC_HeaderInfo）は Collection 代表ドキュメント（README.md）にのみ出力すること
- **FR-027**: 個別 Document の YAML frontmatter には Tbl:DOC_DocumentList の行プロパティのみを含め、Collection ヘッダーの重複を排除すること

### 3.4 FlowBase 共通仕様

- **FR-030**: テーブルには `Tbl:<Prefix>_<TableName>` のスタートマーカーを付与すること（接頭辞付き、ブック内一意）
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

IMG の DocumentProperty.xlsm で定義された全プロパティ。
DocumentBase のテーブルは全列を保持し、role に応じて出力をフィルタする。

### DOC_DocumentList プロパティ

DocumentBase v1.2.0 で DOC_DocumentList に保持するプロパティ。
doctrine_level / phase は doc_type から一意に導出可能なため削除。
domain は Collection 単位で管理（collection_domain）するため削除。

| # | property_name | data_type | input_method | note |
|---|---------------|-----------|-------------|------|
| 1 | no | number | auto | Refresh で自動採番 |
| 2 | title | string | prompt | 行の存在判定 |
| 3 | doc_type | select | suggester | DEF_DocType |
| 4 | doc_type_prefix | string | auto | DEF_DocType.id_prefix から導出 |
| 5 | status | select | suggester | DEF_DocumentStatus（draft/review/active/done/archived） |
| 6 | role | select | suggester | |
| 7 | owner_primary | select | suggester | DEF_Owner（ドキュメント担当者） |
| 8 | version | string | auto | |
| 9 | created | date | auto | |
| 10 | updated | date | auto | |
| 11 | tags | list | manual | |
| 12 | summary | string | manual | |
| 13 | document_id | string | auto | 導出値 |

### DOC_HeaderInfo プロパティ（collection\_ 接頭辞）

| property_name | data_type | note |
|---------------|-----------|------|
| collection_id | string | シート名と同一 |
| collection_name | string | |
| collection_summary | string | |
| collection_domain | select | DEF_CollectionDomain |
| collection_related_project | string | FlowBase 参照（任意） |
| collection_status | select | active / done / archived |
| collection_created | date | |
| collection_updated | date | |
| collection_output_path | path | 個別出力先（任意） |

### IMG プロパティとの対応（参考）

| IMG property | DocumentBase 配置 | 変更理由 |
|---|---|---|
| doctrine_level | DEF_DocType.doctrine_level 列（定義情報として保持） | doc_type から一意導出。DocumentList 列としては冗長 |
| phase | DEF_DocType.phase 列（定義情報として保持） | doc_type から一意導出 |
| domain | DOC_HeaderInfo.collection_domain | コレクション単位で一意 |
| owner_primary | DOC_DocumentList.owner_primary | Collection 管理者→ドキュメント担当者へ移動 |

---

## 7. 開発フェーズ

### Phase 1: IMG でプロパティ定義（完了）

- DocumentProperty.xlsm による M_\* シート定義
- generate_property_master.py による初版生成

### Phase 2: DocumentBase 開発

- Excel ブック作成（M_\* シート + Collection シート）
- VBA: DEF_PropertyMaster 自動集約
- VBA: Templater スクリプト自動生成
- VBA: Obsidian YAML 出力

### Phase 3: IMG に逆導入・再定義

- DocumentBase でプロパティ定義自体を管理（セルフホスティング）
- IMG の Master/ を DocumentBase 出力物に置換
