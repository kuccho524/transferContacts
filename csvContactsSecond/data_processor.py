# データ抽出・突合処理モジュール

from collections import defaultdict
from typing import Any, DefaultDict, Dict, List, Set, Tuple


def normalize_email(value: str) -> str:
    """メールアドレスを比較用に正規化する。"""
    return value.strip().lower()


def unique_preserve_order(values: List[str]) -> List[str]:
    """順序を維持したまま重複を除去する。"""
    seen = set()
    unique_values = []

    for value in values:
        if value and value not in seen:
            seen.add(value)
            unique_values.append(value)

    return unique_values


def names_match(first_name: str, last_name: str, contact: Dict[str, Any]) -> bool:
    """氏名の一致を判定する。"""
    return (
        contact['first_name'] == first_name.strip()
        and contact['last_name'] == last_name.strip()
    )


def build_identity_keys(last_name: str, first_name: str, emails: List[str]) -> List[Tuple[str, str, str]]:
    """氏名とメール群から照合キー一覧を生成する。"""
    return [
        (last_name.strip(), first_name.strip(), normalize_email(email))
        for email in emails
        if normalize_email(email)
    ]


def extract_export_emails(row: Dict[str, str]) -> List[str]:
    """Google Contacts エクスポート行から全メールアドレスを抽出する。"""
    emails = []
    index = 1

    while f'E-mail {index} - Label' in row or f'E-mail {index} - Value' in row:
        value = row.get(f'E-mail {index} - Value', '')
        if value:
            for item in value.split(':::'):
                email = normalize_email(item)
                if email:
                    emails.append(email)
        index += 1

    return unique_preserve_order(emails)


def extract_registered_emails(row: Dict[str, str]) -> List[str]:
    """登録用CSV行から全メールアドレスを抽出する。"""
    emails = []
    for column_name, value in row.items():
        if column_name.endswith('EmailAddress') and value:
            email = normalize_email(value)
            if email:
                emails.append(email)

    return unique_preserve_order(emails)


def extract_contact_emails(row: Dict[str, str]) -> List[str]:
    """contacts.csv 行から全メールアドレスを抽出する。"""
    emails = []
    for column_name, value in row.items():
        if column_name.startswith('emailAddresses.') and column_name.endswith('.value') and value:
            email = normalize_email(value)
            if email:
                emails.append(email)

    return unique_preserve_order(emails)


def find_matching_contact_candidates(
    export_item: Dict[str, Any],
    contact_lookup: Dict[str, Any]
) -> List[Dict[str, Any]]:
    """エクスポート行に対応する候補連絡先を返す。"""
    candidates = {}

    for key in build_identity_keys(
        export_item['last_name'],
        export_item['first_name'],
        export_item['emails']
    ):
        for contact in contact_lookup['by_identity'].get(key, []):
            candidates[contact['resource_name']] = contact

    if candidates:
        return list(candidates.values())

    # 氏名が空のデータに備え、メール一致かつ候補が一意なら採用する。
    for email in export_item['emails']:
        email_matches = contact_lookup['by_email'].get(email, [])
        if len(email_matches) == 1:
            contact = email_matches[0]
            if not export_item['first_name'] and not export_item['last_name']:
                return [contact]

    return []


def extract_export_data(export_data: List[Dict[str, str]]) -> List[Dict[str, Any]]:
    """
    エクスポートデータから必要な列を抽出

    Args:
        export_data: エクスポートデータ

    Returns:
        抽出済みデータ
    """
    extracted = []
    for index, row in enumerate(export_data, start=2):
        emails = extract_export_emails(row)
        extracted.append({
            'row_index': index,
            'first_name': row.get('First Name', '').strip(),
            'last_name': row.get('Last Name', '').strip(),
            'email': emails[0] if emails else '',
            'emails': emails,
            'labels': row.get('Labels', '').strip()
        })

    print(f'✅ エクスポートデータ: {len(extracted)}件抽出')
    return extracted


def extract_registered_data(registered_data: List[Dict[str, str]]) -> List[Dict[str, Any]]:
    """登録用CSVから照合用データを抽出する。"""
    extracted = []

    for row in registered_data:
        emails = extract_registered_emails(row)
        extracted.append({
            'first_name': row.get('FirstName', '').strip(),
            'last_name': row.get('LastName', '').strip(),
            'email': emails[0] if emails else '',
            'emails': emails
        })

    return extracted


def extract_contacts_data(contacts_data: List[Dict[str, str]]) -> List[Dict[str, Any]]:
    """
    連絡先データから必要な列を抽出（people/のみ）

    Args:
        contacts_data: 連絡先データ

    Returns:
        抽出済みデータ
    """
    extracted = []
    people_count = 0
    other_count = 0

    for row in contacts_data:
        resource_name = row.get('resourceName', '').strip()

        if not resource_name.startswith('people/'):
            other_count += 1
            continue

        people_count += 1
        emails = extract_contact_emails(row)

        extracted.append({
            'user': row.get('User', '').strip(),
            'first_name': row.get('names.0.givenName', '').strip(),
            'last_name': row.get('names.0.familyName', '').strip(),
            'email': emails[0] if emails else '',
            'emails': emails,
            'resource_name': resource_name
        })

    print(f'✅ 連絡先データ: {people_count}件抽出 (people/のみ)')
    if other_count > 0:
        print(f'  - otherContacts/は{other_count}件除外')
    return extracted


def extract_labels_data(labels_data: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """
    ラベルデータから必要な列を抽出

    Args:
        labels_data: ラベルデータ

    Returns:
        抽出済みデータ
    """
    extracted = []
    for row in labels_data:
        extracted.append({
            'user': row.get('User', '').strip(),
            'resource_name': row.get('resourceName', '').strip(),
            'name': row.get('name', '').strip()
        })

    print(f'✅ ラベルデータ: {len(extracted)}件抽出')
    return extracted


def match_data(export_data: List[Dict[str, Any]],
               contacts_data: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, str]]]:
    """
    エクスポートデータと連絡先データを突合

    Args:
        export_data: エクスポートデータ（抽出済み）
        contacts_data: 連絡先データ（抽出済み）

    Returns:
        (matched_data, skip_list)
    """
    contact_lookup = {
        'contacts': contacts_data,
        'by_identity': defaultdict(list),
        'by_email': defaultdict(list)
    }

    for contact in contacts_data:
        for key in build_identity_keys(contact['last_name'], contact['first_name'], contact['emails']):
            contact_lookup['by_identity'][key].append(contact)
        for email in contact['emails']:
            contact_lookup['by_email'][email].append(contact)

    matched = []
    skip_list = []

    for export_item in export_data:
        if not export_item['emails']:
            skip_list.append({
                '姓': export_item['last_name'],
                '名': export_item['first_name'],
                'メールアドレス': '(空)',
                '理由': 'メールアドレスが空'
            })
            continue

        candidates = find_matching_contact_candidates(export_item, contact_lookup)

        if len(candidates) == 1:
            matched.append({
                'export': export_item,
                'contact': candidates[0]
            })
        elif len(candidates) > 1:
            skip_list.append({
                '姓': export_item['last_name'],
                '名': export_item['first_name'],
                'メールアドレス': export_item['email'],
                '理由': '一致候補が複数存在'
            })
        else:
            skip_list.append({
                '姓': export_item['last_name'],
                '名': export_item['first_name'],
                'メールアドレス': export_item['email'],
                '理由': '連絡先データに該当なし'
            })

    print(f'✅ 一致: {len(matched)}件')
    if skip_list:
        print(f'⚠️ 不一致: {len(skip_list)}件')

    return matched, skip_list


def parse_labels(labels_str: str) -> List[str]:
    """
    Labels列をパースする

    Args:
        labels_str: ラベル文字列

    Returns:
        ラベルのリスト
    """
    if not labels_str:
        return []

    labels = [label.strip() for label in labels_str.split(':::')]
    return [label for label in labels if label]


def create_label_map(labels_data: List[Dict[str, str]]) -> Dict[str, str]:
    """
    ラベル名からラベルIDへのマッピングを作成

    Args:
        labels_data: ラベルデータ（生のCSVデータ）

    Returns:
        {ラベル名: ラベルID} の辞書
    """
    extracted = extract_labels_data(labels_data)
    label_map = {}

    for label in extracted:
        if label['name'] and label['resource_name']:
            label_map[label['name']] = label['resource_name']

    print(f'✅ ラベルマッピング: {len(label_map)}件作成')
    return label_map


def classify_labels(labels_list: List[str], label_map: Dict[str, str]) -> Dict[str, Any]:
    """
    ラベルをPrimaryとSecondaryに分類

    Args:
        labels_list: ラベルのリスト
        label_map: ラベル名 -> ラベルID のマップ

    Returns:
        {
            'primary_label': str,
            'secondary_labels': List[str]
        }
    """
    primary_label = ''
    secondary_labels = []

    for label_name in labels_list:
        clean_label_name = label_name[2:].strip() if label_name.startswith('* ') else label_name.strip()

        if clean_label_name == 'myContacts':
            primary_label = label_map.get('myContacts', 'contactGroups/myContacts')
        else:
            label_id = label_map.get(clean_label_name, '')
            if label_id:
                secondary_labels.append(label_id)
            else:
                print(f'⚠️ ラベル「{clean_label_name}」が見つかりません（元の名前: {label_name}）')

    return {
        'primary_label': primary_label,
        'secondary_labels': secondary_labels
    }


def get_ordinal_adjective(num: int) -> str:
    """
    数字から序数形容詞を取得

    Args:
        num: 数字（2以上）

    Returns:
        序数形容詞（Third, Fourth, Fifth...）
    """
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

    if num in ordinals:
        return ordinals[num]

    if num % 10 == 1 and num % 100 != 11:
        return f'{num}st'
    if num % 10 == 2 and num % 100 != 12:
        return f'{num}nd'
    if num % 10 == 3 and num % 100 != 13:
        return f'{num}rd'
    return f'{num}th'


def transform_to_output(matched_data: List[Dict[str, Any]],
                        label_map: Dict[str, str]) -> Dict[str, Any]:
    """
    マッチングデータを出力形式に変換

    Args:
        matched_data: マッチング済みデータ
        label_map: ラベルマップ

    Returns:
        {
            'data': List[Dict],
            'max_label_count': int
        }
    """
    output_data = []
    max_label_count = 1

    for item in matched_data:
        export_item = item['export']
        contact_item = item['contact']
        classified = classify_labels(parse_labels(export_item['labels']), label_map)

        output_row = {
            'TargetEmail': contact_item.get('user', ''),
            'ContactID': contact_item['resource_name'],
            'PrimaryLabel': classified['primary_label']
        }

        for idx, label_id in enumerate(classified['secondary_labels']):
            column_name = 'SecondaryLabel' if idx == 0 else f'{get_ordinal_adjective(idx + 1)}Label'
            output_row[column_name] = label_id

        total_labels = 1 + len(classified['secondary_labels'])
        max_label_count = max(max_label_count, total_labels)
        output_data.append(output_row)

    print(f'✅ PrimaryLabel: {len(output_data)}件')
    secondary_count = sum(1 for row in output_data if 'SecondaryLabel' in row)
    if secondary_count > 0:
        print(f'✅ SecondaryLabel: {secondary_count}件')

    return {
        'data': output_data,
        'max_label_count': max_label_count
    }


def generate_output_columns(max_label_count: int) -> List[str]:
    """
    出力CSV列を動的に生成

    Args:
        max_label_count: 最大ラベル数

    Returns:
        列名のリスト
    """
    columns = ['TargetEmail', 'ContactID', 'PrimaryLabel']

    if max_label_count >= 2:
        columns.append('SecondaryLabel')

    for i in range(3, max_label_count + 1):
        columns.append(f'{get_ordinal_adjective(i - 1)}Label')

    return columns


def create_contact_map(contacts_data: List[Dict[str, str]]) -> Dict[str, Any]:
    """
    連絡先データから照合用マップを作成する。

    Args:
        contacts_data: 連絡先データ

    Returns:
        照合用マップ
    """
    extracted = extract_contacts_data(contacts_data)
    by_identity: DefaultDict[Tuple[str, str, str], List[Dict[str, Any]]] = defaultdict(list)
    by_email: DefaultDict[str, List[Dict[str, Any]]] = defaultdict(list)

    for contact in extracted:
        for key in build_identity_keys(contact['last_name'], contact['first_name'], contact['emails']):
            by_identity[key].append(contact)
        for email in contact['emails']:
            by_email[email].append(contact)

    print(f'✅ 連絡先マッピング: {len(extracted)}件作成')
    return {
        'contacts': extracted,
        'by_identity': by_identity,
        'by_email': by_email
    }


def check_consistency(export_data: List[Dict[str, str]],
                      registered_data: List[Dict[str, str]]) -> Dict[str, Any]:
    """
    エクスポートデータと登録データの整合性をチェック

    Args:
        export_data: エクスポートデータ
        registered_data: 登録データ

    Returns:
        {
            'skip_list': List[Dict],
            'skip_row_indexes': Set[int]
        }
    """
    export_extracted = extract_export_data(export_data)
    registered_extracted = extract_registered_data(registered_data)

    registered_identity_set = set()
    for item in registered_extracted:
        for key in build_identity_keys(item['last_name'], item['first_name'], item['emails']):
            registered_identity_set.add(key)

    skip_list = []
    skip_row_indexes: Set[int] = set()

    for export_item in export_extracted:
        if not export_item['emails']:
            skip_list.append({
                '姓': export_item['last_name'],
                '名': export_item['first_name'],
                'メールアドレス': '(空)',
                '対象行': export_item['row_index'],
                '理由': 'メールアドレスが空'
            })
            skip_row_indexes.add(export_item['row_index'])
            continue

        keys = build_identity_keys(
            export_item['last_name'],
            export_item['first_name'],
            export_item['emails']
        )

        if not any(key in registered_identity_set for key in keys):
            skip_list.append({
                '姓': export_item['last_name'],
                '名': export_item['first_name'],
                'メールアドレス': export_item['email'],
                '対象行': export_item['row_index'],
                '理由': '登録データに該当なし'
            })
            skip_row_indexes.add(export_item['row_index'])

    if skip_list:
        print(f'⚠️ 不一致データ: {len(skip_list)}件')
    else:
        print('✅ すべてのデータが一致しています')

    return {
        'skip_list': skip_list,
        'skip_row_indexes': skip_row_indexes
    }


def transform_to_output_data(export_data: List[Dict[str, str]],
                             registered_data: List[Dict[str, str]],
                             contact_map: Dict[str, Any],
                             label_map: Dict[str, str],
                             skip_row_indexes: Set[int],
                             target_email: str) -> Dict[str, Any]:
    """
    エクスポートデータを出力形式に変換（ラベル数ごとにグループ化）

    Args:
        export_data: エクスポートデータ
        registered_data: 登録データ（互換性のため保持）
        contact_map: 連絡先マップ
        label_map: ラベルマップ
        skip_row_indexes: スキップ対象の元データ行番号
        target_email: ターゲットメールアドレス

    Returns:
        {
            'grouped_data': Dict[int, List[Dict]],
            'max_label_count': int
        }
    """
    del registered_data

    export_extracted = extract_export_data(export_data)
    grouped_data: Dict[int, List[Dict[str, str]]] = {}
    max_label_count = 1

    for export_item in export_extracted:
        if export_item['row_index'] in skip_row_indexes or not export_item['emails']:
            continue

        candidates = find_matching_contact_candidates(export_item, contact_map)
        if len(candidates) != 1:
            if len(candidates) > 1:
                print(
                    f"⚠️ 行{export_item['row_index']} は一致候補が複数のためスキップします: "
                    f"{export_item['last_name']} {export_item['first_name']}"
                )
            continue

        contact = candidates[0]
        classified = classify_labels(parse_labels(export_item['labels']), label_map)

        output_row = {
            'TargetEmail': target_email,
            'ContactID': contact['resource_name'],
            'PrimaryLabel': classified['primary_label']
        }

        for idx, label_id in enumerate(classified['secondary_labels']):
            column_name = 'SecondaryLabel' if idx == 0 else f'{get_ordinal_adjective(idx + 1)}Label'
            output_row[column_name] = label_id

        total_labels = 1 + len(classified['secondary_labels'])
        max_label_count = max(max_label_count, total_labels)
        grouped_data.setdefault(total_labels, []).append(output_row)

    total_count = sum(len(data) for data in grouped_data.values())
    print(f'✅ 出力データ: {total_count}件生成')
    for label_count in sorted(grouped_data.keys()):
        print(f'  - ラベル数{label_count}: {len(grouped_data[label_count])}件')

    return {
        'grouped_data': grouped_data,
        'max_label_count': max_label_count
    }
