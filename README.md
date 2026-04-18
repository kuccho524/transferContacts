# Google Contacts 連絡先移行ツール

Google Contacts のエクスポート CSV をもとに、連絡先登録用 CSV の生成、GAM を使った登録、登録後データの照合、ラベル付与用 CSV の生成までを補助する Windows 向けツール群です。

このリポジトリには実行プログラム、補助スクリプト、設計書、匿名化サンプルを含みます。実行時に生成される `runs/` と `logs/` は公開対象外です。

## 公開版の前提条件

- Windows 10 以降
- Python 3.6 以降
- PowerShell 5.1 以降
- [GAM](https://github.com/GAM-team/GAM) の導入と認証が完了していること
- Google Workspace 側で必要な権限があること
- 日本語ラベルを扱う場合は Git Bash の導入を推奨

## できること

1. Google Contacts エクスポート CSV から連絡先登録用 CSV とラベル作成用 CSV を生成する
2. GAM 経由で連絡先を登録する
3. 登録済みの contacts / contactgroups データを取得する
4. 元データと登録後データを突合し、ラベル付与用 CSV を生成する
5. 統合 GUI から一連の処理を順に実行する

## 推奨導線

公開版での基本導線は次の 2 つです。

- 統合GUI: `python transferContacts_master.py`
- Step 6 単体CLI: `python csvContactsSecond/main.py ...`

公開版では現行導線のみを含め、旧単体GUIや旧補助スクリプトは含めません。

## 起動方法

リポジトリの `transferContacts` ディレクトリで実行します。

```bash
python transferContacts_master.py
```

## 主な入力

- 旧アカウントからエクスポートした Google Contacts CSV
- 移行先アカウントのメールアドレス

## 主な出力

- 連絡先登録用 CSV
- ラベル作成用 CSV
- GAM から取得した `contacts.csv`
- GAM から取得した `contactgroups.csv`
- ラベル登録用 `contacts_labels_*.csv`

これらの生成物は実行時に `runs/` や `logs/` に作られますが、公開リポジトリには含めない運用を想定しています。

## 基本的な処理フロー

1. エクスポート CSV を選択する
2. Step 2 で登録用 CSV を生成する
3. Step 3 で連絡先を登録する
4. Step 4 でラベルを作成する
5. Step 5 で登録後データを取得する
6. Step 6 で突合してラベル登録用 CSV を生成する
7. Step 7 でラベルを付与する

## ディレクトリ構成

```text
transferContacts/
├── transferContacts_master.py
├── csvContactsFirst/
│   └── convert_contacts.py
├── csvContactsSecond/
│   ├── main.py
│   ├── csv_parser.py
│   ├── data_processor.py
│   ├── csv_generator.py
│   └── README.md
├── registerContacts/
│   ├── invoke_gam.ps1
│   ├── run_register_cli.ps1
├── registerLabels/
│   └── registerLabelsToContacts.txt
├── samples/
│   ├── export_contacts.csv
│   ├── registered_contacts.csv
│   ├── contacts.csv
│   ├── contactgroups.csv
│   └── README.md
├── tests/
│   ├── test_convert_contacts_ordinals.py
│   └── test_public_samples.py
└── docs/
```

## 運用上の注意

- このツールは Windows と GAM の利用を前提としています
- 実行時に扱う CSV やログには個人情報が含まれる可能性があります
- `runs/`, `logs/`, `__pycache__/` は Git に含めないでください
- 公開前に README、設計書、サンプルファイルに実データが残っていないことを確認してください

## ドキュメント

- 統合版の設計資料は `docs/` を参照してください
- Step 6 相当の単体ツール説明は `csvContactsSecond/README.md` を参照してください
- 匿名化済み入出力例は `samples/` を参照してください
- 公開版では `legacy/` 配下の旧導線は配布対象に含めません
- 変数名・関数名は英語だが、その説明コメントは日本語
- console.log()のメッセージは日本語
- エラーメッセージは日本語で出力

### 設計書

詳細な設計情報は以下のドキュメントを参照してください：

- [基本設計書](docs/基本設計書_統合版.md)
- [詳細設計書](docs/詳細設計書_統合版.md)

---

## ライセンス

[MIT License](LICENSE)

---

## 作成者

Claude Sonnet 4.5

---

## バージョン履歴

| バージョン | 日付 | 変更内容 |
|----------|------|---------|
| 1.0 | 2025-12-30 | 初版リリース |

---

## サポート

問題が発生した場合は、以下の情報を含めて報告してください：

- エラーメッセージ（スクリーンショット）
- 実行ログ（`./logs/` 内のログファイル）
- 実行環境（OS、Pythonバージョン、GAMバージョン）

## テスト

匿名化サンプルを使った最小確認は次で実行できます。

```bash
python -m unittest discover -s tests
python csvContactsSecond/test_ordinal.py
```
