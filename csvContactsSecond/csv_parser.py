# CSV読み込み・パース処理モジュール

import csv
from typing import List, Dict


def parse_csv(file_path: str) -> List[Dict[str, str]]:
    """
    CSVファイルを読み込んでパースする

    Args:
        file_path: 読み込むファイルパス

    Returns:
        パース済みデータの配列（辞書のリスト）

    Raises:
        FileNotFoundError: ファイルが見つからない場合
        UnicodeDecodeError: UTF-8以外のエンコーディングの場合
        ValueError: CSV形式が不正な場合
    """
    try:
        # UTF-8で読み込み
        with open(file_path, 'r', encoding='utf-8', newline='') as f:
            # BOM付きUTF-8の場合は除去
            content = f.read()
            if content.startswith('\ufeff'):
                content = content[1:]

            # エンコーディング検証
            if not validate_encoding(content):
                raise ValueError(f'ファイル {file_path} のエンコーディングがUTF-8ではありません。UTF-8形式のファイルを用意してください。')

            # CSV読み込み
            lines = content.splitlines()
            reader = csv.DictReader(lines)

            data = []
            for row in reader:
                data.append(row)

            if len(data) == 0:
                print(f'⚠️ {file_path}: データが空です')

            return data

    except UnicodeDecodeError:
        raise ValueError(f'ファイル {file_path} のエンコーディングがUTF-8ではありません。UTF-8形式のファイルを用意してください。')
    except FileNotFoundError:
        raise FileNotFoundError(f'ファイル {file_path} が見つかりません。')
    except Exception as e:
        raise ValueError(f'ファイル {file_path} の読み込みに失敗しました: {str(e)}')


def validate_encoding(text: str) -> bool:
    """
    エンコーディングを検証

    Args:
        text: 検証対象テキスト

    Returns:
        UTF-8として正しいか
    """
    # 置換文字（U+FFFD）の検出
    if '\ufffd' in text:
        return False

    # 複数の置換文字が連続するパターン
    if '��' in text:
        return False

    # 制御文字（タブ、改行以外）
    for char in text:
        code = ord(char)
        if 0x00 <= code <= 0x08 or code == 0x0B or code == 0x0C or 0x0E <= code <= 0x1F:
            return False

    return True
