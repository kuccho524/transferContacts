#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
連絡先ラベル登録CSV生成ツール（Python版）

Google Contactsのエクスポートデータと登録済み連絡先データを突合し、
ラベル登録用のCSVファイルを生成する
"""

import os
import sys

# Windows環境でのUTF-8出力を有効化（インポート前に設定）
if sys.platform == 'win32':
    import io
    if hasattr(sys.stdout, 'buffer'):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    if hasattr(sys.stderr, 'buffer'):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    # 環境変数も設定
    os.environ['PYTHONIOENCODING'] = 'utf-8'

import argparse
import re
from typing import Dict, List

from csv_parser import parse_csv
from data_processor import (
    create_label_map,
    create_contact_map,
    check_consistency,
    transform_to_output_data,
    generate_output_columns
)
from csv_generator import (
    generate_csv,
    save_csv,
    generate_filename,
    save_skip_list_csv
)


def validate_email(email: str) -> bool:
    """
    メールアドレス形式を検証

    Args:
        email: メールアドレス

    Returns:
        有効な形式の場合True
    """
    pattern = r'^[^\s@]+@[^\s@]+\.[^\s@]+$'
    return re.match(pattern, email) is not None


def main():
    """メイン処理"""
    # コマンドライン引数のパース
    parser = argparse.ArgumentParser(
        description='連絡先ラベル登録CSV生成ツール',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用例:
  python main.py \\
    --export-data テスト用エクスポートデータ.csv \\
    --registered-data 連絡先登録データ_サンプル.csv \\
    --label-data contactgroups.csv \\
    --contacts-data contacts.csv \\
    --target-email user@example.com \\
    --output contacts_labels.csv
        '''
    )

    parser.add_argument(
        '--export-data',
        required=True,
        help='エクスポート元データのCSVファイルパス'
    )
    parser.add_argument(
        '--registered-data',
        required=True,
        help='登録後データのCSVファイルパス'
    )
    parser.add_argument(
        '--label-data',
        required=True,
        help='ラベルデータ（contactgroups.csv）のファイルパス'
    )
    parser.add_argument(
        '--contacts-data',
        required=True,
        help='登録済み連絡先データ（contacts.csv）のファイルパス'
    )
    parser.add_argument(
        '--target-email',
        required=True,
        help='登録対象アカウントのメールアドレス'
    )
    parser.add_argument(
        '--output',
        default=None,
        help='出力CSVファイルパス（省略時は自動生成）'
    )
    parser.add_argument(
        '--skip-list-output',
        default=None,
        help='スキップリストCSVファイルパス（省略時は自動生成）'
    )

    args = parser.parse_args()

    try:
        print('=' * 60)
        print('連絡先ラベル登録CSV生成ツール')
        print('=' * 60)
        print('')

        # ターゲットメールアドレスの検証
        if not validate_email(args.target_email):
            raise ValueError('メールアドレスの形式が正しくありません')

        print(f'✅ 登録対象アカウント: {args.target_email}')
        print('')

        # Phase 1: ファイル読み込み
        print('=== Phase 1: ファイル読み込み ===')

        export_data = parse_csv(args.export_data)
        print(f'✅ エクスポート元データ: {len(export_data)}件読み込み完了')

        registered_data = parse_csv(args.registered_data)
        print(f'✅ 登録後データ: {len(registered_data)}件読み込み完了')

        label_data = parse_csv(args.label_data)
        print(f'✅ ラベルデータ: {len(label_data)}件読み込み完了')

        contacts_data = parse_csv(args.contacts_data)
        print(f'✅ 登録済み連絡先データ: {len(contacts_data)}件読み込み完了')
        print('')

        # Phase 2: データマッピング
        print('=== Phase 2: データマッピング ===')
        label_map = create_label_map(label_data)
        contact_map = create_contact_map(contacts_data)
        print('')

        # Phase 3: 整合性チェック
        print('=== Phase 3: 整合性チェック ===')
        consistency_result = check_consistency(export_data, registered_data)

        # スキップリスト表示
        if consistency_result['skip_list']:
            print('')
            print('⚠️ スキップされた連絡先:')
            for item in consistency_result['skip_list']:
                print(f"  - {item['姓']} {item['名']} ({item['メールアドレス']}) - 行{item['対象行']}")

            # スキップリストをCSV出力
            if args.skip_list_output:
                skip_list_path = args.skip_list_output
            else:
                skip_list_path = generate_filename('skip_list')

            save_skip_list_csv(consistency_result['skip_list'], skip_list_path)
        else:
            print('✅ すべてのデータが一致しています')
        print('')

        # Phase 4: データ突合・変換
        print('=== Phase 4: データ突合・変換 ===')
        skip_row_indexes = consistency_result['skip_row_indexes']
        transform_result = transform_to_output_data(
            export_data,
            registered_data,
            contact_map,
            label_map,
            skip_row_indexes,
            args.target_email
        )
        print('')

        # Phase 5: CSV出力（ラベル数ごとに分割）
        print('=== Phase 5: CSV出力 ===')
        grouped_data = transform_result['grouped_data']

        # 出力ファイルのベースパス決定
        import os
        if args.output:
            base_dir = os.path.dirname(args.output)
            base_name = os.path.splitext(os.path.basename(args.output))[0]
            # base_dirが空の場合はカレントディレクトリ
            if not base_dir:
                base_dir = '.'
        else:
            base_dir = '.'
            base_name = 'contacts_labels'

        # ラベル数ごとにCSVファイルを生成
        output_files = []
        for label_count in sorted(grouped_data.keys()):
            data = grouped_data[label_count]

            # 列名を生成
            columns = generate_output_columns(label_count)

            # CSVコンテンツを生成
            csv_content = generate_csv(data, columns)

            # 出力ファイルパス
            output_filename = f'{base_name}_{label_count}.csv'
            output_path = os.path.join(base_dir, output_filename)

            save_csv(csv_content, output_path)
            output_files.append(output_path)

        print('')
        print(f'✅ 合計{len(output_files)}個のCSVファイルを生成しました')
        for output_file in output_files:
            print(f'  - {output_file}')
        print('')

        print('✨ すべての処理が完了しました！')
        print('=' * 60)

    except FileNotFoundError as e:
        print(f'❌ エラー: {str(e)}')
        sys.exit(1)
    except ValueError as e:
        print(f'❌ エラー: {str(e)}')
        sys.exit(1)
    except Exception as e:
        print(f'❌ 予期しないエラーが発生しました: {str(e)}')
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
