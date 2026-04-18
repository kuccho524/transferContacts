# 連絡先ラベル登録 CSV 生成ツール

Google Contacts のエクスポートデータ、登録用 CSV、`contactgroups.csv`、`contacts.csv` を突合し、ラベル付与用の `contacts_labels_*.csv` を生成する補助ツールです。

このディレクトリの実行導線は CLI 版です。

## 前提条件

- Python 3.6 以降
- 入力 CSV は UTF-8
- `contacts.csv` は GAM などから取得した Google Contacts データ
- `contactgroups.csv` は GAM などから取得した連絡先グループデータ

## 実行方法

### CLI 版

```bash
python main.py \
  --export-data <export.csv> \
  --registered-data <registered.csv> \
  --label-data <contactgroups.csv> \
  --contacts-data <contacts.csv> \
  --target-email user@example.com
```

## 引数

| 引数 | 必須 | 説明 |
|------|------|------|
| `--export-data` | ✓ | 元のエクスポート CSV |
| `--registered-data` | ✓ | Step 2 で生成した登録用 CSV |
| `--label-data` | ✓ | ラベル一覧 CSV |
| `--contacts-data` | ✓ | 登録済み連絡先一覧 CSV |
| `--target-email` | ✓ | ラベル付与対象アカウント |
| `--output` | - | 出力ファイルのベース名 |
| `--skip-list-output` | - | スキップリスト CSV の出力先 |

## 入力の想定

### エクスポート CSV

- `First Name`
- `Last Name`
- `E-mail 1 - Value`
- `Labels`

### ラベル一覧 CSV

- `resourceName`
- `name`

### 連絡先一覧 CSV

- `resourceName`
- `emailAddresses.0.value`
- `names.0.givenName`
- `names.0.familyName`

`resourceName` が `people/` で始まる連絡先を主対象として扱います。

## 出力

- `contacts_labels_1.csv`
- `contacts_labels_2.csv`
- `contacts_labels_3.csv`
- `skip_list_*.csv`

ラベル数ごとにファイルを分割して出力します。

## ファイル構成

```text
csvContactsSecond/
├── main.py
├── csv_parser.py
├── data_processor.py
├── csv_generator.py
├── test_ordinal.py
└── docs/
```

## 注意事項

- 入力データに個人情報が含まれる可能性があります
- 実データや生成物は公開リポジトリに含めないでください
- `contacts.csv` と `contactgroups.csv` の取得方法は上位 README と設計書を参照してください
- 匿名化された検証用データは `../samples/` を参照してください

## テスト

```bash
python test_ordinal.py
python -m unittest ..\\tests\\test_public_samples.py
```
