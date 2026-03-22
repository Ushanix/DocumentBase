# DocumentBase

**DocumentBase** は、ドキュメントのプロパティ管理と Obsidian 連携のための Excel ベースツールです。

InformationManagementGuidelines で定義されたドキュメントプロパティ規格に準拠し、
Excel マスタシートから Obsidian Templater スクリプトを自動生成します。

---

## Repository Structure

```
DocumentBase/
├── docs/            設計書・仕様
│   └── design/      要件定義・設計ドキュメント
├── excel/           DocumentBase-Excel (VBA)
│   ├── DocumentBase.xlsm
│   └── src/         エクスポート済み VBA モジュール
└── README.md
```

---

## Concept

### Base シリーズにおける位置づけ

| Base | 対象 | 時間依存 |
|------|------|----------|
| **FlowBase** | プロジェクト・タスク管理 | 年度依存 |
| **TimeBase** | 日報・計画・ライフログ | 日次〜年次 |
| **DocumentBase** | ドキュメント・ナレッジ管理 | 時間非依存 |

DocumentBase は時空に依存しない知識資産を管理するためのツールです。
FlowBase がワークフロードリブン、TimeBase が時間ドリブンであるのに対し、
DocumentBase はマスタドリブンで文書のメタデータを構造化します。

---

## Related Repositories

| Repository | Description |
|---|---|
| [InformationManagementGuidelines](../InformationManagementGuidelines) | ドキュメントプロパティ規格・マスタ定義 |
| [FlowBase](../FlowBase) | プロジェクト・タスク管理 |
| [TimeBase](../TimeBase) | 日報・計画・ライフログ |

---

## License

MIT License
