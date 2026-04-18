#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSVデータ突合ツール
エクスポートデータから連絡先登録用CSVとラベル作成用CSVを生成する
"""

import csv
import argparse
import sys
import logging
from pathlib import Path
from datetime import datetime


def setup_logging(log_level='INFO'):
    """
    ログ設定を初期化

    Args:
        log_level: ログレベル（'DEBUG', 'INFO', 'WARNING', 'ERROR'）

    Returns:
        logging.Logger: 設定済みのロガー
    """
    # logsディレクトリを作成
    log_dir = Path('logs')
    log_dir.mkdir(exist_ok=True)

    # ログファイル名を生成
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = log_dir / f'convert_{timestamp}.log'

    # ログフォーマット
    log_format = '%(asctime)s - %(levelname)s - %(message)s'
    date_format = '%Y-%m-%d %H:%M:%S'

    # ロガーを取得
    logger = logging.getLogger('csv_converter')
    logger.setLevel(getattr(logging, log_level.upper()))

    # 既存のハンドラをクリア（重複を防ぐ）
    logger.handlers.clear()

    # ファイルハンドラ（UTF-8で保存）
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(logging.Formatter(log_format, date_format))

    # コンソールハンドラ
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter(log_format, date_format))

    # ハンドラを追加
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logger.info(f'ログファイル: {log_file}')
    logger.info(f'ログレベル: {log_level.upper()}')

    return logger


def select_files():
    """
    ファイル選択ダイアログを表示してファイルパスを取得

    Returns:
        tuple: (input_file, contact_output, label_output)
               キャンセルされた場合は None
    """
    import tkinter as tk
    from tkinter import filedialog

    # ルートウィンドウを作成（非表示）
    root = tk.Tk()
    root.withdraw()

    # 入力ファイルを選択
    input_file = filedialog.askopenfilename(
        title="エクスポートCSVファイルを選択",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )

    if not input_file:
        print("キャンセルされました")
        return None

    # 連絡先CSV出力先を選択
    contact_output = filedialog.asksaveasfilename(
        title="連絡先CSVの保存先を選択",
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        initialfile="連絡先.csv"
    )

    if not contact_output:
        print("キャンセルされました")
        return None

    # ラベルCSV出力先を選択
    label_output = filedialog.asksaveasfilename(
        title="ラベルCSVの保存先を選択",
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        initialfile="ラベル.csv"
    )

    if not label_output:
        print("キャンセルされました")
        return None

    # ルートウィンドウを破棄
    root.destroy()

    return input_file, contact_output, label_output


def read_export_data(input_file, logger=None):
    """
    エクスポートデータを読み込む

    Args:
        input_file: 入力ファイルパス
        logger: ロガーインスタンス

    Returns:
        list: レコードのリスト
    """
    records = []
    try:
        if logger:
            logger.debug(f'ファイル読み込み開始: {input_file}')

        with open(input_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                records.append(row)

        if logger:
            logger.debug(f'ファイル読み込み完了: {len(records)}件')

    except Exception as e:
        if logger:
            logger.error(f'ファイルの読み込みに失敗: {e}', exc_info=True)
        print(f"エラー: ファイルの読み込みに失敗しました: {e}", file=sys.stderr)
        sys.exit(1)

    return records


def get_ordinal_suffix(num):
    """
    序数の接尾辞を取得する (Primary, Secondary, Third, Fourth, ...)

    Args:
        num: 番号 (1始まり)

    Returns:
        str: 序数接尾辞
    """
    suffixes = {
        1: 'Primary',
        2: 'Secondary',
        3: 'Third',
        4: 'Fourth',
        5: 'Fifth',
        6: 'Sixth',
        7: 'Seventh',
        8: 'Eighth',
        9: 'Ninth',
        10: 'Tenth'
    }
    return suffixes.get(num, f'{num}th')


def extract_emails_and_phones(row):
    """
    レコードからメールアドレスと電話番号を抽出する
    " ::: " で区切られた複数の値がある場合は、それぞれを個別のエントリとして扱う

    Args:
        row: CSVレコード

    Returns:
        tuple: (emails, phones) のリスト
    """
    emails = []
    phones = []

    # メールアドレスを抽出
    i = 1
    while f'E-mail {i} - Label' in row or f'E-mail {i} - Value' in row:
        label = row.get(f'E-mail {i} - Label', '')
        value = row.get(f'E-mail {i} - Value', '')
        if value:  # 値がある場合のみ追加
            # " ::: " で区切られた複数の値がある場合は分割
            values = [v.strip() for v in value.split(' ::: ') if v.strip()]
            for val in values:
                emails.append({'label': label, 'value': val})
        i += 1

    # 電話番号を抽出
    i = 1
    while f'Phone {i} - Label' in row or f'Phone {i} - Value' in row:
        label = row.get(f'Phone {i} - Label', '')
        value = row.get(f'Phone {i} - Value', '')
        if value:  # 値がある場合のみ追加
            # " ::: " で区切られた複数の値がある場合は分割
            values = [v.strip() for v in value.split(' ::: ') if v.strip()]
            for val in values:
                phones.append({'label': label, 'value': val})
        i += 1

    return emails, phones


def generate_contact_csv(data, output_file, logger=None):
    """
    連絡先登録用CSVを生成

    Args:
        data: レコードのリスト
        output_file: 出力ファイルパス
        logger: ロガーインスタンス
    """
    if logger:
        logger.debug('連絡先CSV生成開始')

    # 最大のメール・電話件数を調査
    max_emails = 0
    max_phones = 0

    processed_data = []
    for row in data:
        emails, phones = extract_emails_and_phones(row)
        max_emails = max(max_emails, len(emails))
        max_phones = max(max_phones, len(phones))
        processed_data.append({
            'row': row,
            'emails': emails,
            'phones': phones
        })

    if logger:
        logger.debug(f'最大メール件数: {max_emails}, 最大電話件数: {max_phones}')

    # ヘッダー行を生成
    headers = [
        'FirstName',
        'LastName',
        'HiraganaFirstName',
        'HiraganaLastName',
        'OrganizationName',
        'OrganizationDepartment'
    ]

    # メールアドレスのカラムを追加
    for i in range(1, max_emails + 1):
        suffix = get_ordinal_suffix(i)
        headers.append(f'{suffix}EmailLabel')
        headers.append(f'{suffix}EmailAddress')

    # 電話番号のカラムを追加
    for i in range(1, max_phones + 1):
        suffix = get_ordinal_suffix(i)
        headers.append(f'{suffix}PhoneLabel')
        headers.append(f'{suffix}PhoneNumber')

    # CSV出力
    try:
        with open(output_file, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f, lineterminator='\n')
            writer.writerow(headers)

            for item in processed_data:
                row = item['row']
                emails = item['emails']
                phones = item['phones']

                record = [
                    row.get('First Name', ''),
                    row.get('Last Name', ''),
                    row.get('Phonetic First Name', ''),
                    row.get('Phonetic Last Name', ''),
                    row.get('Organization Name', ''),
                    row.get('Organization Department', '')
                ]

                # メールアドレスを追加
                for i in range(max_emails):
                    if i < len(emails):
                        record.append(emails[i]['label'])
                        record.append(emails[i]['value'])
                    else:
                        record.append('')
                        record.append('')

                # 電話番号を追加
                for i in range(max_phones):
                    if i < len(phones):
                        record.append(phones[i]['label'])
                        record.append(phones[i]['value'])
                    else:
                        record.append('')
                        record.append('')

                writer.writerow(record)

        if logger:
            logger.info(f'連絡先CSVを生成しました: {output_file}')
        print(f"連絡先CSVを生成しました: {output_file}")

    except Exception as e:
        if logger:
            logger.error(f'連絡先CSVの生成に失敗: {e}', exc_info=True)
        print(f"エラー: 連絡先CSVの生成に失敗しました: {e}", file=sys.stderr)
        sys.exit(1)


def parse_labels(labels_str):
    """
    ラベル文字列を分割してすべてのラベルを返す
    ":::" で分割し、リストで返す（スペースの有無に関係なく対応）

    Args:
        labels_str: ラベル文字列

    Returns:
        list: ラベルのリスト（空の場合は空リスト）
    """
    if not labels_str:
        return []

    # ":::" で分割（スペースがあっても trim() で処理）
    parts = labels_str.split(':::')

    # トリムして空でない要素のみ返す
    labels = [part.strip() for part in parts if part.strip()]

    return labels


def get_label_column_name(index):
    """
    ラベル列の列名を取得（PrimaryLabel, SecondaryLabel, ThirdLabel, FourthLabel...）

    Args:
        index: インデックス（0始まり）

    Returns:
        str: 列名
    """
    ordinals = {
        0: 'PrimaryLabel',
        1: 'SecondaryLabel',
        2: 'ThirdLabel',
        3: 'FourthLabel',
        4: 'FifthLabel',
        5: 'SixthLabel',
        6: 'SeventhLabel',
        7: 'EighthLabel',
        8: 'NinthLabel',
        9: 'TenthLabel',
        10: 'EleventhLabel',
        11: 'TwelfthLabel',
        12: 'ThirteenthLabel',
        13: 'FourteenthLabel',
        14: 'FifteenthLabel',
        15: 'SixteenthLabel',
        16: 'SeventeenthLabel',
        17: 'NineteenthLabel',
        18: 'TwentiethLabel',
        19: 'TwentyFirstLabel',
        20: 'TwentySecondLabel',
        21: 'TwentyThirdLabel',
    }

    if index in ordinals:
        return ordinals[index]
    else:
        # 22以降は数値表記
        num = index + 1  # PrimaryLabelが0なので+1
        if num % 10 == 1 and num % 100 != 11:
            return f'{num}stLabel'
        elif num % 10 == 2 and num % 100 != 12:
            return f'{num}ndLabel'
        elif num % 10 == 3 and num % 100 != 13:
            return f'{num}rdLabel'
        else:
            return f'{num}thLabel'


def generate_label_csv(data, output_file, logger=None):
    """
    ラベル作成用CSVを生成（動的に3つ以上のラベル列に対応）

    PrimaryLabel: 必ず "* myContacts"
    SecondaryLabel以降: "* myContacts" 以外のラベル

    Args:
        data: レコードのリスト
        output_file: 出力ファイルパス
        logger: ロガーインスタンス
    """
    if logger:
        logger.debug('ラベルCSV生成開始')

    try:
        # すべてのレコードからラベルを抽出し、最大ラベル数を調査
        max_secondary_label_count = 0
        all_labels = []

        for row in data:
            labels_str = row.get('Labels', '')
            labels = parse_labels(labels_str)

            # "* myContacts"以外のラベルを抽出
            secondary_labels = []
            for label in labels:
                # "* myContacts"または"myContacts"は除外
                clean_label = label.strip()
                if clean_label == '* myContacts' or clean_label == 'myContacts':
                    continue
                secondary_labels.append(clean_label)

            all_labels.append(secondary_labels)
            max_secondary_label_count = max(max_secondary_label_count, len(secondary_labels))

        if logger:
            logger.debug(f'最大SecondaryLabel数: {max_secondary_label_count}')

        # ヘッダー行を動的に生成（PrimaryLabel, SecondaryLabel, ThirdLabel, FourthLabel...）
        # 最大列数 = 1 (PrimaryLabel) + max_secondary_label_count
        max_total_columns = 1 + max_secondary_label_count
        headers = []
        for i in range(max_total_columns):
            headers.append(get_label_column_name(i))

        if logger:
            logger.debug(f'生成される列: {headers}')

        # CSV出力
        with open(output_file, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f, lineterminator='\n')
            writer.writerow(headers)

            for secondary_labels in all_labels:
                row_data = []

                # PrimaryLabel: 必ず "* myContacts"
                row_data.append('* myContacts')

                # SecondaryLabel以降
                for i in range(max_secondary_label_count):
                    if i < len(secondary_labels):
                        row_data.append(secondary_labels[i])
                    else:
                        row_data.append('')

                writer.writerow(row_data)

        if logger:
            logger.info(f'ラベルCSVを生成しました: {output_file}（PrimaryLabel + Secondary {max_secondary_label_count}列）')
        print(f"ラベルCSVを生成しました: {output_file}（PrimaryLabel + Secondary {max_secondary_label_count}列）")

    except Exception as e:
        if logger:
            logger.error(f'ラベルCSVの生成に失敗: {e}', exc_info=True)
        print(f"エラー: ラベルCSVの生成に失敗しました: {e}", file=sys.stderr)
        sys.exit(1)


def main():
    """メイン処理"""
    parser = argparse.ArgumentParser(
        description='CSVデータ突合ツール - エクスポートデータから連絡先CSVとラベルCSVを生成'
    )
    parser.add_argument('input_file', nargs='?', help='入力ファイル（エクスポートデータのCSVファイルパス）')
    parser.add_argument('contact_output', nargs='?', help='連絡先CSV出力ファイル名')
    parser.add_argument('label_output', nargs='?', help='ラベルCSV出力ファイル名')
    parser.add_argument('--log-level', default='INFO',
                       choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
                       help='ログレベル（デフォルト: INFO）')

    args = parser.parse_args()

    # ログ設定を初期化
    logger = setup_logging(args.log_level)
    logger.info('=' * 60)
    logger.info('CSVデータ突合ツール 処理開始')
    logger.info('=' * 60)

    # コマンドライン引数が指定されているかチェック
    if args.input_file and args.contact_output and args.label_output:
        # コマンドライン引数を使用
        input_file = args.input_file
        contact_output = args.contact_output
        label_output = args.label_output
        logger.info('コマンドライン引数からファイルパスを取得')
    else:
        # ファイル選択ダイアログを使用
        logger.info('ファイル選択ダイアログを表示')
        result = select_files()
        if result is None:
            logger.info('ユーザーによる処理キャンセル')
            print("処理を中止しました")
            sys.exit(0)
        input_file, contact_output, label_output = result
        logger.info('ファイル選択ダイアログからファイルパスを取得')

    # ファイルパスをログに記録
    logger.info(f'入力ファイル: {input_file}')
    logger.info(f'連絡先CSV出力: {contact_output}')
    logger.info(f'ラベルCSV出力: {label_output}')

    # 入力ファイルの存在確認
    if not Path(input_file).exists():
        logger.error(f'入力ファイルが見つかりません: {input_file}')
        print(f"エラー: 入力ファイルが見つかりません: {input_file}", file=sys.stderr)
        sys.exit(1)

    # データ読み込み
    logger.info('データ読み込み開始')
    data = read_export_data(input_file, logger)
    logger.info(f'{len(data)}件のレコードを読み込みました')

    # CSV生成
    logger.info('CSV生成開始')
    try:
        generate_contact_csv(data, contact_output, logger)
        logger.info('連絡先CSVの生成完了')

        generate_label_csv(data, label_output, logger)
        logger.info('ラベルCSVの生成完了')
    except Exception as e:
        logger.error(f'CSV生成中にエラーが発生: {e}', exc_info=True)
        raise

    # 処理完了
    logger.info('=' * 60)
    logger.info(f'処理完了: {len(data)}件のレコードを処理しました')
    logger.info('=' * 60)
    print(f"処理完了: {len(data)}件のレコードを処理しました")


if __name__ == '__main__':
    main()
