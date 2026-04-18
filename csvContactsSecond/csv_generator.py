# CSV出力処理モジュール

from datetime import datetime
from typing import List, Dict


def generate_csv(data: List[Dict[str, str]], columns: List[str]) -> str:
    """
    データからCSV文字列を生成

    Args:
        data: 出力データ
        columns: 列名のリスト

    Returns:
        CSV文字列（UTF-8、LF改行）
    """
    lines = []

    # ヘッダー行
    lines.append(','.join(columns))

    # データ行
    for row in data:
        fields = []
        for column in columns:
            value = row.get(column, '')
            if value is None:
                value = ''
            fields.append(escape_csv_field(str(value)))
        lines.append(','.join(fields))

    # LF改行で結合
    return '\n'.join(lines)


def escape_csv_field(value: str) -> str:
    """
    CSVフィールドをエスケープ

    Args:
        value: フィールド値

    Returns:
        エスケープ済みフィールド
    """
    # ダブルクォート、カンマ、改行が含まれる場合はダブルクォートで囲む
    if '"' in value or ',' in value or '\n' in value or '\r' in value:
        # ダブルクォートを2つに
        escaped_value = value.replace('"', '""')
        return f'"{escaped_value}"'
    else:
        return value


def save_csv(csv_content: str, output_path: str) -> None:
    """
    CSV文字列をファイルに保存

    Args:
        csv_content: CSV文字列
        output_path: 出力ファイルパス
    """
    # UTF-8、BOM無し、LF改行で保存
    with open(output_path, 'w', encoding='utf-8', newline='') as f:
        f.write(csv_content)

    print(f'✅ CSV出力完了: {output_path}')


def generate_filename(prefix: str = 'contacts_labels') -> str:
    """
    タイムスタンプ付きファイル名を生成

    Args:
        prefix: ファイル名のプレフィックス

    Returns:
        ファイル名（例: contacts_labels_20231228_143022.csv）
    """
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return f'{prefix}_{timestamp}.csv'


def generate_skip_list_csv(skip_list: List[Dict[str, str]]) -> str:
    """
    スキップリストのCSV文字列を生成

    Args:
        skip_list: スキップリスト

    Returns:
        CSV文字列
    """
    if not skip_list:
        return ''

    columns = ['姓', '名', 'メールアドレス', '理由']
    lines = []

    # ヘッダー行
    lines.append(','.join(columns))

    # データ行
    for item in skip_list:
        fields = [
            escape_csv_field(item.get('姓', '')),
            escape_csv_field(item.get('名', '')),
            escape_csv_field(item.get('メールアドレス', '')),
            escape_csv_field(item.get('理由', ''))
        ]
        lines.append(','.join(fields))

    return '\n'.join(lines)


def save_skip_list_csv(skip_list: List[Dict[str, str]], output_path: str) -> None:
    """
    スキップリストをCSVファイルに保存

    Args:
        skip_list: スキップリスト
        output_path: 出力ファイルパス
    """
    csv_content = generate_skip_list_csv(skip_list)

    if csv_content:
        with open(output_path, 'w', encoding='utf-8', newline='') as f:
            f.write(csv_content)
        print(f'✅ スキップリストCSV出力完了: {output_path}')
    else:
        print('⚠️ スキップリストが空です')
