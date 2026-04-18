#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
序数形容詞生成関数のテスト
"""

import sys
import os

# Windows環境でのUTF-8出力を有効化
if sys.platform == 'win32':
    import io
    if hasattr(sys.stdout, 'buffer'):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    if hasattr(sys.stderr, 'buffer'):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

sys.path.insert(0, '.')

from data_processor import get_ordinal_adjective

def test_ordinal_adjective():
    """序数形容詞生成のテスト"""
    print("=== 序数形容詞生成テスト ===\n")

    # テストケース
    # 注意: この関数は「Primary（1番目）」「Secondary（2番目）」の次の序数形容詞を返す
    # 引数2 → Third（3番目）、引数3 → Fourth（4番目）
    test_cases = [
        (2, 'Third'),
        (3, 'Fourth'),
        (4, 'Fifth'),
        (10, 'Eleventh'),
        (19, 'Twentieth'),
        (20, 'TwentyFirst'),
        (25, 'TwentySixth'),
        (30, 'ThirtyFirst'),
        (40, 'FortyFirst'),
        (49, 'Fiftieth'),
        (50, '50th'),  # 51以降は数値表記
        (51, '51st'),
        (52, '52nd'),
        (53, '53rd'),
        (60, '60th'),
        (100, '100th'),
        (111, '111th'),
        (112, '112th'),
    ]

    print("テストケース実行:")
    all_passed = True

    for num, expected in test_cases:
        result = get_ordinal_adjective(num)
        status = "✅" if result == expected else "❌"
        if result != expected:
            all_passed = False
        print(f"{status} {num} → {result} (期待値: {expected})")

    print("\n" + "=" * 50)

    # 実際の列名生成例
    print("\n=== 実際の列名生成例 ===\n")
    print("3個のラベルがある場合:")
    for i in range(1, 4):
        if i == 1:
            print(f"  {i}個目: PrimaryLabel")
        elif i == 2:
            print(f"  {i}個目: SecondaryLabel")
        else:
            print(f"  {i}個目: {get_ordinal_adjective(i - 1)}Label")

    print("\n5個のラベルがある場合:")
    for i in range(1, 6):
        if i == 1:
            print(f"  {i}個目: PrimaryLabel")
        elif i == 2:
            print(f"  {i}個目: SecondaryLabel")
        else:
            print(f"  {i}個目: {get_ordinal_adjective(i - 1)}Label")

    print("\n10個のラベルがある場合:")
    for i in range(1, 11):
        if i == 1:
            print(f"  {i}個目: PrimaryLabel")
        elif i == 2:
            print(f"  {i}個目: SecondaryLabel")
        else:
            print(f"  {i}個目: {get_ordinal_adjective(i - 1)}Label")

    print("\n55個のラベルがある場合（抜粋）:")
    for i in [1, 2, 3, 10, 20, 30, 50, 51, 52, 53, 55]:
        if i == 1:
            print(f"  {i}個目: PrimaryLabel")
        elif i == 2:
            print(f"  {i}個目: SecondaryLabel")
        else:
            print(f"  {i}個目: {get_ordinal_adjective(i - 1)}Label")

    print("\n" + "=" * 50)

    if all_passed:
        print("\n✅ すべてのテストが成功しました！")
    else:
        print("\n❌ 一部のテストが失敗しました")

    return all_passed

if __name__ == '__main__':
    success = test_ordinal_adjective()
    sys.exit(0 if success else 1)
