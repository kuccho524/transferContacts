# 詳細設計書
## 連絡先ラベル登録CSV生成ツール（Python版）

> 注記: 現行の公開導線は CLI 版 `csvContactsSecond/main.py` です。旧単体 GUI は公開版に含めず、現在の挙動はソースコードを優先してください。

---

## Phase 1: ファイル読み込み

### 目的
3つのCSVファイルを読み込み、UTF-8エンコーディングを検証する

### 入力
- 連絡先エクスポートデータのファイルパス
- 登録した連絡先のデータのファイルパス
- 登録したラベルのデータのファイルパス

### 出力
- `export_data`: List[Dict[str, str]] - エクスポートデータ
- `contacts_data`: List[Dict[str, str]] - 連絡先データ
- `labels_data`: List[Dict[str, str]] - ラベルデータ

### 処理フロー
```python
def load_files(export_path, contacts_path, labels_path):
    # 1. エクスポートデータ読み込み
    export_data = parse_csv(export_path)
    log(f"✅ 連絡先エクスポートデータ: {len(export_data)}件読み込み完了")

    # 2. 連絡先データ読み込み
    contacts_data = parse_csv(contacts_path)
    log(f"✅ 登録した連絡先のデータ: {len(contacts_data)}件読み込み完了")

    # 3. ラベルデータ読み込み
    labels_data = parse_csv(labels_path)
    log(f"✅ 登録したラベルのデータ: {len(labels_data)}件読み込み完了")

    return export_data, contacts_data, labels_data
```

### エンコーディング検証
```python
def validate_encoding(text: str) -> bool:
    # 置換文字（U+FFFD）の検出
    if '\ufffd' in text:
        return False

    # 制御文字チェック（タブ、改行以外）
    for char in text:
        code = ord(char)
        if 0x00 <= code <= 0x08 or code == 0x0B or code == 0x0C or 0x0E <= code <= 0x1F:
            return False

    return True
```

### エラー処理
- **FileNotFoundError**: ファイルが見つからない
- **UnicodeDecodeError**: UTF-8以外のエンコーディング
- **ValueError**: CSV形式が不正

---

## Phase 2: データ抽出

### 目的
各ファイルから必要な列のみを抽出し、フィルタリングする

### 2.1 エクスポートデータ抽出

**抽出列**:
- `First Name`: 名
- `Last Name`: 姓
- `E-mail 1 - Value`: メールアドレス
- `Labels`: ラベル情報

**処理**:
```python
def extract_export_data(export_data):
    extracted = []
    for row in export_data:
        extracted.append({
            'first_name': row.get('First Name', '').strip(),
            'last_name': row.get('Last Name', '').strip(),
            'email': row.get('E-mail 1 - Value', '').strip(),
            'labels': row.get('Labels', '').strip()
        })

    log(f"✅ エクスポートデータ: {len(extracted)}件抽出")
    return extracted
```

### 2.2 連絡先データ抽出

**抽出列**:
- `User`: ユーザーメールアドレス
- `names.0.givenName`: 名
- `names.0.familyName`: 姓
- `emailAddresses.0.value`: メールアドレス
- `resourceName`: 連絡先ID

**フィルタ条件**: `resourceName`が`people/`で始まるもののみ

**処理**:
```python
def extract_contacts_data(contacts_data):
    extracted = []
    people_count = 0
    other_count = 0

    for row in contacts_data:
        resource_name = row.get('resourceName', '').strip()

        # people/のみ処理
        if not resource_name.startswith('people/'):
            other_count += 1
            continue

        people_count += 1

        extracted.append({
            'user': row.get('User', '').strip(),
            'first_name': row.get('names.0.givenName', '').strip(),
            'last_name': row.get('names.0.familyName', '').strip(),
            'email': row.get('emailAddresses.0.value', '').strip(),
            'resource_name': resource_name
        })

    log(f"✅ 連絡先データ: {people_count}件抽出 (people/のみ)")
    log(f"  - otherContacts/は{other_count}件除外")
    return extracted
```

### 2.3 ラベルデータ抽出

**抽出列**:
- `User`: ユーザーメールアドレス
- `resourceName`: ラベルID
- `name`: ラベル名

**処理**:
```python
def extract_labels_data(labels_data):
    extracted = []
    for row in labels_data:
        extracted.append({
            'user': row.get('User', '').strip(),
            'resource_name': row.get('resourceName', '').strip(),
            'name': row.get('name', '').strip()
        })

    log(f"✅ ラベルデータ: {len(extracted)}件抽出")
    return extracted
```

---

## Phase 3: データ突合

### 目的
エクスポートデータと連絡先データを突合し、一致するデータを抽出する

### 一致判定ルール
**条件**: 以下の3項目がすべて一致
1. `First Name` (エクスポート) = `names.0.givenName` (連絡先)
2. `Last Name` (エクスポート) = `names.0.familyName` (連絡先)
3. `E-mail 1 - Value` (エクスポート) = `emailAddresses.0.value` (連絡先)

**正規化**:
- メールアドレス: 小文字に変換
- 姓名: トリム後の完全一致

### 処理フロー
```python
def match_data(export_data, contacts_data):
    # 連絡先データをマップ化（高速検索用）
    contacts_map = {}
    for contact in contacts_data:
        # キー: (姓, 名, メールアドレス小文字)
        key = (
            contact['last_name'],
            contact['first_name'],
            contact['email'].lower()
        )
        contacts_map[key] = contact

    matched = []
    skip_list = []

    for export_item in export_data:
        # 検索キー作成
        key = (
            export_item['last_name'],
            export_item['first_name'],
            export_item['email'].lower()
        )

        # 一致判定
        if key in contacts_map:
            contact = contacts_map[key]
            matched.append({
                'export': export_item,
                'contact': contact
            })
        else:
            # 不一致 → スキップリストへ
            skip_list.append({
                '姓': export_item['last_name'],
                '名': export_item['first_name'],
                'メールアドレス': export_item['email'],
                '理由': '連絡先データに該当なし'
            })

    log(f"✅ 一致: {len(matched)}件")
    if skip_list:
        log(f"⚠️ 不一致: {len(skip_list)}件")

    return matched, skip_list
```

---

## Phase 4: ラベル解析

### 目的
エクスポートデータのLabels列を解析し、ラベルIDに変換する

### 4.1 Labels列のパース

**入力**: `Team A ::: * myContacts`
**出力**: `['Team A', '* myContacts']`

**処理**:
```python
def parse_labels(labels_str: str) -> List[str]:
    if not labels_str:
        return []

    # ":::"で分割
    labels = [label.strip() for label in labels_str.split(':::')]

    # 空文字列を除外
    labels = [label for label in labels if label]

    return labels
```

### 4.2 ラベルマッピング作成

**処理**:
```python
def create_label_map(labels_data):
    label_map = {}

    for label in labels_data:
        name = label['name']
        resource_name = label['resource_name']

        if name and resource_name:
            label_map[name] = resource_name

    log(f"✅ ラベルマッピング: {len(label_map)}件作成")
    return label_map
```

### 4.3 ラベル分類

**ルール**:
- `* myContacts` → PrimaryLabel
- その他 → SecondaryLabel, ThirdLabel, FourthLabel...（動的に生成）

**動的対応**:
- 入力データに含まれるラベル数を自動検出
- 必要な列数を動的に生成（制限なし）
- 3個、4個、それ以上のラベルにも対応

**処理**:
```python
def classify_labels(labels_list, label_map):
    primary_label = ''
    secondary_labels = []

    for label_name in labels_list:
        # * myContactsの場合
        if label_name == '* myContacts':
            # label_mapから'myContacts'を検索
            primary_label = label_map.get('myContacts', 'contactGroups/myContacts')
        else:
            # ユーザー定義ラベル（個数制限なし）
            label_id = label_map.get(label_name, '')
            if label_id:
                secondary_labels.append(label_id)
            else:
                log(f"⚠️ ラベル「{label_name}」が見つかりません")

    return {
        'primary_label': primary_label,
        'secondary_labels': secondary_labels  # 3個以上のラベルにも対応
    }
```

### 4.4 出力データ生成

**処理**:
```python
def transform_to_output(matched_data, label_map):
    output_data = []
    max_label_count = 1  # PrimaryLabel

    for item in matched_data:
        export_item = item['export']
        contact_item = item['contact']

        # ラベル解析
        labels_list = parse_labels(export_item['labels'])
        classified = classify_labels(labels_list, label_map)

        # 出力行作成
        output_row = {
            'TargetEmail': contact_item['user'],
            'ContactID': contact_item['resource_name'],
            'PrimaryLabel': classified['primary_label']
        }

        # SecondaryLabel以降を追加
        for idx, label_id in enumerate(classified['secondary_labels']):
            if idx == 0:
                column_name = 'SecondaryLabel'
            else:
                column_name = f'{get_ordinal_adjective(idx + 2)}Label'

            output_row[column_name] = label_id

        # 最大ラベル数を更新
        total_labels = 1 + len(classified['secondary_labels'])
        if total_labels > max_label_count:
            max_label_count = total_labels

        output_data.append(output_row)

    log(f"✅ PrimaryLabel: {len(output_data)}件")
    log(f"✅ SecondaryLabel: {sum(1 for row in output_data if 'SecondaryLabel' in row)}件")

    return {
        'data': output_data,
        'max_label_count': max_label_count
    }
```

### 4.5 序数形容詞生成

**処理**:
```python
def get_ordinal_adjective(num: int) -> str:
    """
    数字から序数形容詞を生成

    Args:
        num: 数字（2以上）

    Returns:
        序数形容詞（Third, Fourth, Fifth...）

    Examples:
        2 → Third
        3 → Fourth
        4 → Fifth
        20 → Twentieth
        21 → TwentyFirst
        50 → Fiftieth
        51 → 51st
        52 → 52nd
        53 → 53rd
        54 → 54th
    """
    # 基本的な序数形容詞（2-50）
    ordinals = {
        2: 'Third',
        3: 'Fourth',
        4: 'Fifth',
        5: 'Sixth',
        6: 'Seventh',
        7: 'Eighth',
        8: 'Ninth',
        9: 'Tenth',
        10: 'Eleventh',
        11: 'Twelfth',
        12: 'Thirteenth',
        13: 'Fourteenth',
        14: 'Fifteenth',
        15: 'Sixteenth',
        16: 'Seventeenth',
        17: 'Eighteenth',
        18: 'Nineteenth',
        19: 'Twentieth',
        20: 'TwentyFirst',
        21: 'TwentySecond',
        22: 'TwentyThird',
        23: 'TwentyFourth',
        24: 'TwentyFifth',
        25: 'TwentySixth',
        26: 'TwentySeventh',
        27: 'TwentyEighth',
        28: 'TwentyNinth',
        29: 'Thirtieth',
        30: 'ThirtyFirst',
        31: 'ThirtySecond',
        32: 'ThirtyThird',
        33: 'ThirtyFourth',
        34: 'ThirtyFifth',
        35: 'ThirtySixth',
        36: 'ThirtySeventh',
        37: 'ThirtyEighth',
        38: 'ThirtyNinth',
        39: 'Fortieth',
        40: 'FortyFirst',
        41: 'FortySecond',
        42: 'FortyThird',
        43: 'FortyFourth',
        44: 'FortyFifth',
        45: 'FortySixth',
        46: 'FortySeventh',
        47: 'FortyEighth',
        48: 'FortyNinth',
        49: 'Fiftieth'
    }

    # 辞書にある場合はそれを返す
    if num in ordinals:
        return ordinals[num]

    # 51以降は数値表記（51st, 52nd, 53rd, 54th...）
    if num % 10 == 1 and num % 100 != 11:
        return f'{num}st'
    elif num % 10 == 2 and num % 100 != 12:
        return f'{num}nd'
    elif num % 10 == 3 and num % 100 != 13:
        return f'{num}rd'
    else:
        return f'{num}th'
```

**仕様**:
- 2-50の数字には英単語の序数形容詞を使用（Third, Fourth...Fiftieth）
- 51以降は数値表記（51st, 52nd, 53rd, 54th...）
- 実質的に無制限のラベル数に対応可能

---

## Phase 5: CSV出力

### 目的
変換済みデータをCSV形式で出力する

### 5.1 列名生成

**処理**:
```python
def generate_output_columns(max_label_count):
    columns = ['TargetEmail', 'ContactID', 'PrimaryLabel']

    if max_label_count >= 2:
        columns.append('SecondaryLabel')

    for i in range(3, max_label_count + 1):
        columns.append(f'{get_ordinal_adjective(i - 1)}Label')

    return columns
```

### 5.2 CSV生成

**処理**:
```python
def generate_csv(data, columns):
    lines = []

    # ヘッダー行
    lines.append(','.join(columns))

    # データ行
    for row in data:
        fields = []
        for column in columns:
            value = row.get(column, '')
            fields.append(escape_csv_field(str(value)))
        lines.append(','.join(fields))

    # LF改行で結合
    return '\n'.join(lines)
```

### 5.3 CSVフィールドエスケープ

**処理**:
```python
def escape_csv_field(value: str) -> str:
    # ダブルクォート、カンマ、改行が含まれる場合はエスケープ
    if '"' in value or ',' in value or '\n' in value or '\r' in value:
        escaped_value = value.replace('"', '""')
        return f'"{escaped_value}"'
    else:
        return value
```

### 5.4 ファイル保存

**処理**:
```python
def save_csv(csv_content, output_path):
    # UTF-8、BOM無し、LF改行で保存
    with open(output_path, 'w', encoding='utf-8', newline='') as f:
        f.write(csv_content)

    log(f"✅ CSV出力完了: {output_path}")
```

### 5.5 スキップリスト出力

**処理**:
```python
def save_skip_list(skip_list, output_path):
    if not skip_list:
        return

    columns = ['姓', '名', 'メールアドレス', '理由']
    lines = [','.join(columns)]

    for item in skip_list:
        fields = [
            escape_csv_field(item['姓']),
            escape_csv_field(item['名']),
            escape_csv_field(item['メールアドレス']),
            escape_csv_field(item['理由'])
        ]
        lines.append(','.join(fields))

    csv_content = '\n'.join(lines)

    with open(output_path, 'w', encoding='utf-8', newline='') as f:
        f.write(csv_content)

    log(f"✅ スキップリストCSV出力: {output_path}")
```

---

## データ構造定義

### エクスポートデータ
```python
{
    'first_name': str,      # 名
    'last_name': str,       # 姓
    'email': str,           # メールアドレス
    'labels': str           # ラベル（:::区切り）
}
```

### 連絡先データ
```python
{
    'user': str,            # ユーザーメール
    'first_name': str,      # 名
    'last_name': str,       # 姓
    'email': str,           # メールアドレス
    'resource_name': str    # people/XXXXXXXXXX
}
```

### ラベルデータ
```python
{
    'user': str,            # ユーザーメール
    'resource_name': str,   # contactGroups/XXXXXXXXXX
    'name': str             # ラベル名
}
```

### 出力データ
```python
{
    'TargetEmail': str,     # 対象GWSアカウント
    'ContactID': str,       # people/XXXXXXXXXX
    'PrimaryLabel': str,    # contactGroups/myContacts
    'SecondaryLabel': str,  # contactGroups/XXXXXXXXXX (任意)
    'ThirdLabel': str,      # contactGroups/XXXXXXXXXX (任意)
    ...
}
```

---

**作成日**: 2025-12-28
**バージョン**: 2.0
**作成者**: Claude Code
