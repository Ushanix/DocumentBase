# Templater 自動生成設計

## 1. 概要

M_\* シートのヘッダー情報から Obsidian Templater スクリプトを VBA マクロで機械的に生成する。
AI による質的プロンプトではなく、パターンベースの決定的な生成である。

---

## 2. 生成パターン

M_\* シートの `input_method` と `parent_sheet` により、生成されるコードが一意に決まる。

### 2.1 基本パターン

| input_method | parent_sheet | 生成コード |
|---|---|---|
| prompt | — | `await tp.system.prompt("display_name")` |
| suggester | 空 | `await tp.system.suggester(displays, values, true, "display_name")` |
| suggester | あり | 親の選択結果でフィルタ → suggester |
| auto | — | 固定値または `tp.date.now()` の直接代入 |
| manual | — | 出力のみ（入力プロンプトなし） |

### 2.2 階層選択パターン（parent_sheet あり）

```javascript
// 例: M_Project (parent_sheet: M_ProjectGroup)
// Step 1: 親を選択
const projectGroup = await tp.system.suggester(
  parentDisplays, parentValues, true, "プロジェクト大区分を選択："
);
meta.projectGroup = projectGroup;

// Step 2: 親の値でフィルタして子を選択
const filtered = allProjects.filter(p => p.parent === projectGroup);
const project = await tp.system.suggester(
  filtered.map(p => p.display),
  filtered.map(p => p.value),
  true, "プロジェクトを選択："
);
meta.project = project;
```

---

## 3. 生成フロー

```
1. M_* シートを列挙（シート名が "M_" で始まるもの）
2. 各シートのヘッダー領域（行1-9）を読み取り
3. applies_to でグルーピング:
   - "all" → 共通セクション
   - それ以外 → role 分岐セクション
4. yaml_order でソート
5. テンプレートに流し込み:
   ┌─ 共通関数定義（toYaml 等）         ← 固定文字列
   ├─ 共通プロパティの入力コード生成      ← パターン適用
   ├─ role 選択                         ← M_Role から生成
   ├─ if/else 分岐の生成                ← applies_to から自動構成
   │   ├─ role=project の固有プロパティ
   │   ├─ role=task の固有プロパティ（※FlowBase側で管理）
   │   └─ role=docs/dashboard の固有プロパティ（※doctrine_level は doc_type から導出）
   ├─ 共通後処理（日付・version 等）      ← auto プロパティ
   ├─ YAML order 配列の生成             ← yaml_order 列から自動構成
   └─ toYaml 呼び出し                   ← 固定文字列
6. .md ファイルとして出力
```

---

## 4. M_\* シート追加時の影響

NFR-003 の通り、M_\* シートの追加のみで新プロパティを導入できる。

1. 新しい M_\* シートを追加（ヘッダー領域 + データ領域）
2. DEF_PropertyMaster 自動集約を実行 → 新プロパティが反映
3. Templater 生成を実行 → 新プロパティの入力コードが自動追加

VBA コードの変更は不要。

---

## 5. DocumentBase v1.2.0 でのスキーマ変更の影響

以下のプロパティは DocumentBase の DOC_DocumentList から削除されたが、
Templater 生成には引き続き IMG の M_\* シート定義を使用するため、
Templater スクリプト自体には影響しない。

| プロパティ | 変更内容 | Templater への影響 |
|---|---|---|
| doctrine_level | DOC_DocumentList から削除。DEF_DocType に doctrine_level 列として保持。 | Templater では doc_type 選択時に doctrine_level を自動導出可能。M_DoctrineLevel を事前選択UIとして使う場合は従来通り。 |
| phase | DOC_DocumentList から削除。doc_type から導出。 | Templater 側で必要なら M_Phase を参照。DocumentBase 出力には含まない。 |
| domain | DOC_DocumentList から削除。Collection 単位で管理（collection_domain）。 | Templater で個別ノートに domain を付与する場合は M_Domain を参照。DocumentBase ではコレクション単位。 |
| status | DOC_DocumentList に追加（draft/review/active/done/archived） | Templater の初期値は `draft` を推奨。 |
| owner_primary | DOC_HeaderInfo から DOC_DocumentList に移動 | Templater で入力プロンプトを生成。 |
