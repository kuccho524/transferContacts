#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Google Contacts 連絡先移行ツール - 統合版

旧Googleアカウントから新アカウントへ連絡先を移行する全工程（手順2～7）を
1つのGUIアプリケーションで一気通貫で実行する
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import subprocess
import threading
import queue
import shutil
import json
from datetime import datetime
from pathlib import Path
import re

# Windows環境でのUTF-8出力を有効化
if sys.platform == 'win32':
    import io
    if hasattr(sys.stdout, 'buffer'):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    if hasattr(sys.stderr, 'buffer'):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    os.environ['PYTHONIOENCODING'] = 'utf-8'


class TransferContactsMasterGUI:
    """連絡先移行統合GUIアプリケーション"""

    def __init__(self, root):
        self.root = root
        self.root.title("Google Contacts 連絡先移行ツール - 統合版")
        self.root.geometry("1000x900")

        # スクリプトのベースディレクトリを取得
        self.base_dir = Path(__file__).parent.resolve()
        self.python_executable = sys.executable or "python"
        self.gam_executable = shutil.which("gam")
        self.gam_wrapper_script = self.base_dir / "registerContacts" / "invoke_gam.ps1"
        self.log_queue = queue.Queue()

        # ファイルパス変数
        self.export_csv_path = tk.StringVar()
        self.target_email = tk.StringVar()
        self.contact_csv_path = tk.StringVar()
        self.label_csv_path = tk.StringVar()
        self.log_dir = tk.StringVar(value="./logs")

        # 処理ステップ有効/無効
        self.step2_enabled = tk.BooleanVar(value=True)
        self.step3_enabled = tk.BooleanVar(value=True)
        self.step4_enabled = tk.BooleanVar(value=True)
        self.step5_enabled = tk.BooleanVar(value=True)
        self.step6_enabled = tk.BooleanVar(value=True)
        self.step7_enabled = tk.BooleanVar(value=True)

        # 中間ファイル
        self.work_dir = None
        self.intermediate_contact_csv = None
        self.intermediate_label_csv = None
        self.intermediate_contacts_csv = None
        self.intermediate_contactgroups_csv = None
        self.final_labels_csv_list = []  # 複数のラベル登録用CSVファイル

        # UI構築
        self.create_widgets()
        self.root.after(100, self.flush_log_queue)

        # 初期ログ出力
        self.log(f"Python実行環境: {self.python_executable}")
        if self.gam_executable:
            self.log(f"✅ GAM検出: {self.gam_executable}")
        else:
            self.log("⚠️ GAM が PATH で見つかりません。GAMを使うStep 3以降は実行できません。")

    def create_widgets(self):
        """UI構築"""
        # タイトル
        title_frame = tk.Frame(self.root, bg='#2c3e50', pady=15)
        title_frame.pack(fill=tk.X)

        title_label = tk.Label(
            title_frame,
            text="Google Contacts 連絡先移行ツール - 統合版",
            font=('Yu Gothic UI', 16, 'bold'),
            bg='#2c3e50',
            fg='white'
        )
        title_label.pack()

        # メインフレーム
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ① 入力ファイル選択セクション
        input_section = tk.LabelFrame(
            main_frame,
            text="① 入力ファイル選択",
            padx=15,
            pady=15,
            font=('Yu Gothic UI', 10, 'bold')
        )
        input_section.pack(fill=tk.X, pady=(0, 10))

        # エクスポートCSV
        tk.Label(
            input_section,
            text="エクスポートCSV:",
            width=20,
            anchor=tk.W
        ).grid(row=0, column=0, sticky=tk.W, pady=5)

        tk.Entry(
            input_section,
            textvariable=self.export_csv_path,
            width=50
        ).grid(row=0, column=1, sticky=tk.EW, padx=(0, 10), pady=5)

        tk.Button(
            input_section,
            text="参照...",
            command=lambda: self.browse_file(self.export_csv_path, "エクスポートCSVを選択")
        ).grid(row=0, column=2, pady=5)

        # ターゲットメールアドレス
        tk.Label(
            input_section,
            text="ターゲットメールアドレス:",
            width=20,
            anchor=tk.W
        ).grid(row=1, column=0, sticky=tk.W, pady=5)

        tk.Entry(
            input_section,
            textvariable=self.target_email,
            width=50
        ).grid(row=1, column=1, sticky=tk.EW, pady=5)

        input_section.columnconfigure(1, weight=1)

        # ② 出力先設定セクション
        output_section = tk.LabelFrame(
            main_frame,
            text="② 出力先設定（任意、未指定時は自動生成）",
            padx=15,
            pady=15,
            font=('Yu Gothic UI', 10, 'bold')
        )
        output_section.pack(fill=tk.X, pady=(0, 10))

        # 連絡先CSV出力先
        tk.Label(
            output_section,
            text="連絡先CSV:",
            width=20,
            anchor=tk.W
        ).grid(row=0, column=0, sticky=tk.W, pady=5)

        tk.Entry(
            output_section,
            textvariable=self.contact_csv_path,
            width=50
        ).grid(row=0, column=1, sticky=tk.EW, padx=(0, 10), pady=5)

        tk.Button(
            output_section,
            text="参照...",
            command=lambda: self.browse_save_file(
                self.contact_csv_path,
                "連絡先CSV保存先を選択",
                "連絡先"
            )
        ).grid(row=0, column=2, pady=5)

        # ラベルCSV出力先
        tk.Label(
            output_section,
            text="ラベルCSV:",
            width=20,
            anchor=tk.W
        ).grid(row=1, column=0, sticky=tk.W, pady=5)

        tk.Entry(
            output_section,
            textvariable=self.label_csv_path,
            width=50
        ).grid(row=1, column=1, sticky=tk.EW, padx=(0, 10), pady=5)

        tk.Button(
            output_section,
            text="参照...",
            command=lambda: self.browse_save_file(
                self.label_csv_path,
                "ラベルCSV保存先を選択",
                "ラベル"
            )
        ).grid(row=1, column=2, pady=5)

        # ログ出力先
        tk.Label(
            output_section,
            text="ログ出力先:",
            width=20,
            anchor=tk.W
        ).grid(row=2, column=0, sticky=tk.W, pady=5)

        tk.Entry(
            output_section,
            textvariable=self.log_dir,
            width=50
        ).grid(row=2, column=1, sticky=tk.EW, padx=(0, 10), pady=5)

        tk.Button(
            output_section,
            text="参照...",
            command=lambda: self.browse_directory(self.log_dir, "ログ出力先を選択")
        ).grid(row=2, column=2, pady=5)

        output_section.columnconfigure(1, weight=1)

        # ③ 処理ステップ選択セクション
        step_section = tk.LabelFrame(
            main_frame,
            text="③ 処理ステップ選択",
            padx=15,
            pady=15,
            font=('Yu Gothic UI', 10, 'bold')
        )
        step_section.pack(fill=tk.X, pady=(0, 10))

        tk.Checkbutton(
            step_section,
            text="Step 2: CSV抽出",
            variable=self.step2_enabled,
            font=('Yu Gothic UI', 10)
        ).grid(row=0, column=0, sticky=tk.W, pady=2)

        tk.Checkbutton(
            step_section,
            text="Step 3: 連絡先登録",
            variable=self.step3_enabled,
            font=('Yu Gothic UI', 10)
        ).grid(row=1, column=0, sticky=tk.W, pady=2)

        tk.Checkbutton(
            step_section,
            text="Step 4: ラベル作成",
            variable=self.step4_enabled,
            font=('Yu Gothic UI', 10)
        ).grid(row=2, column=0, sticky=tk.W, pady=2)

        tk.Checkbutton(
            step_section,
            text="Step 5: 登録データ取得",
            variable=self.step5_enabled,
            font=('Yu Gothic UI', 10)
        ).grid(row=0, column=1, sticky=tk.W, pady=2, padx=(30, 0))

        tk.Checkbutton(
            step_section,
            text="Step 6: データ突合",
            variable=self.step6_enabled,
            font=('Yu Gothic UI', 10)
        ).grid(row=1, column=1, sticky=tk.W, pady=2, padx=(30, 0))

        tk.Checkbutton(
            step_section,
            text="Step 7: ラベル登録",
            variable=self.step7_enabled,
            font=('Yu Gothic UI', 10)
        ).grid(row=2, column=1, sticky=tk.W, pady=2, padx=(30, 0))

        # ボタンフレーム
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 10))

        # 設定を確認ボタン
        tk.Button(
            button_frame,
            text="設定を確認",
            command=self.confirm_settings,
            bg='#3498db',
            fg='white',
            font=('Yu Gothic UI', 11, 'bold'),
            padx=20,
            pady=8
        ).pack(side=tk.LEFT, padx=(0, 10))

        # 処理開始ボタン
        self.process_button = tk.Button(
            button_frame,
            text="処理開始 ▶",
            command=self.confirm_and_run,
            bg='#27ae60',
            fg='white',
            font=('Yu Gothic UI', 12, 'bold'),
            padx=30,
            pady=8
        )
        self.process_button.pack(side=tk.LEFT, padx=(0, 10))

        # クリアボタン
        tk.Button(
            button_frame,
            text="クリア",
            command=self.clear_all,
            bg='#e74c3c',
            fg='white',
            font=('Yu Gothic UI', 11),
            padx=20,
            pady=8
        ).pack(side=tk.LEFT)

        # ログ出力エリア
        log_section = tk.LabelFrame(
            main_frame,
            text="実行ログ",
            padx=10,
            pady=10,
            font=('Yu Gothic UI', 10, 'bold')
        )
        log_section.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_section,
            height=20,
            bg='#1e1e1e',
            fg='#00ff00',
            font=('Consolas', 9),
            wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def browse_file(self, path_var, title):
        """ファイル選択ダイアログ"""
        filename = filedialog.askopenfilename(
            title=title,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            path_var.set(filename)
            self.log(f"ファイル選択: {filename}")

    def browse_save_file(self, path_var, title, default_name):
        """保存ファイル選択ダイアログ"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_filename = f"{default_name}_{timestamp}.csv"

        filename = filedialog.asksaveasfilename(
            title=title,
            defaultextension=".csv",
            initialfile=default_filename,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            path_var.set(filename)
            self.log(f"保存先選択: {filename}")

    def browse_directory(self, path_var, title):
        """ディレクトリ選択ダイアログ"""
        dirname = filedialog.askdirectory(title=title)
        if dirname:
            path_var.set(dirname)
            self.log(f"ディレクトリ選択: {dirname}")

    def append_log_line(self, log_line):
        """ログウィジェットに1行追記する。"""
        self.log_text.insert(tk.END, log_line + '\n')
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def flush_log_queue(self):
        """ワーカースレッドからのログをUIスレッドで反映する。"""
        while True:
            try:
                log_line = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.append_log_line(log_line)

        self.root.after(100, self.flush_log_queue)

    def log(self, message):
        """ログメッセージ追加"""
        timestamp = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
        log_line = f"[{timestamp}] {message}"

        if threading.current_thread() is threading.main_thread():
            self.append_log_line(log_line)
        else:
            self.log_queue.put(log_line)

    def clear_all(self):
        """すべてのフィールドをクリア"""
        self.export_csv_path.set('')
        self.target_email.set('')
        self.contact_csv_path.set('')
        self.label_csv_path.set('')
        self.log_dir.set('./logs')

        # チェックボックスを全てON
        self.step2_enabled.set(True)
        self.step3_enabled.set(True)
        self.step4_enabled.set(True)
        self.step5_enabled.set(True)
        self.step6_enabled.set(True)
        self.step7_enabled.set(True)

        # ログクリア
        self.log_text.delete(1.0, tk.END)
        self.log('クリア完了')

    def generate_confirmation_message(self):
        """確認メッセージ生成"""
        enabled_steps = []
        if self.step2_enabled.get():
            enabled_steps.append("✅ Step 2: CSV抽出")
        if self.step3_enabled.get():
            enabled_steps.append("✅ Step 3: 連絡先登録")
        if self.step4_enabled.get():
            enabled_steps.append("✅ Step 4: ラベル作成")
        if self.step5_enabled.get():
            enabled_steps.append("✅ Step 5: 登録データ取得")
        if self.step6_enabled.get():
            enabled_steps.append("✅ Step 6: データ突合")
        if self.step7_enabled.get():
            enabled_steps.append("✅ Step 7: ラベル登録")

        steps_text = "\n  ".join(enabled_steps)

        message = f"""━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  実行設定の確認
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【入力】
  エクスポートCSV: {self.export_csv_path.get() or '(未選択)'}
  ターゲットメール: {self.target_email.get() or '(未入力)'}

【出力先】
  連絡先CSV: {self.contact_csv_path.get() or '(自動生成)'}
  ラベルCSV: {self.label_csv_path.get() or '(自動生成)'}
  ログ: {self.log_dir.get()}

【実行する処理】
  {steps_text}
"""
        return message

    def requires_gam(self):
        """GAM が必要なStepが選択されているか判定する。"""
        return any([
            self.step3_enabled.get(),
            self.step4_enabled.get(),
            self.step5_enabled.get(),
            self.step6_enabled.get(),
            self.step7_enabled.get(),
        ])

    def existing_input_path(self, value):
        """既存入力ファイルとして利用できるパスを返す。"""
        if not value:
            return None

        path = Path(value)
        return path if path.exists() else None

    def validate_step_dependencies(self):
        """選択されたStepの依存関係を検証する。"""
        if self.requires_gam() and not self.gam_executable:
            return "Step 3以降を実行するには、GAM が PATH に含まれている必要があります。"

        existing_contact_csv = self.existing_input_path(self.contact_csv_path.get())
        existing_label_csv = self.existing_input_path(self.label_csv_path.get())

        if self.step3_enabled.get() and not self.step2_enabled.get() and not existing_contact_csv:
            return "Step 3 を単独で実行する場合は、既存の連絡先CSVを指定してください。"

        if self.step4_enabled.get() and not self.step2_enabled.get() and not existing_label_csv:
            return "Step 4 を単独で実行する場合は、既存のラベルCSVを指定してください。"

        if self.step6_enabled.get() and not self.step5_enabled.get():
            return "Step 6 を実行するには Step 5 を有効にしてください。"

        if self.step6_enabled.get() and not self.step2_enabled.get() and not existing_contact_csv:
            return "Step 6 を実行するには、既存の連絡先CSVを指定するか Step 2 を有効にしてください。"

        if self.step7_enabled.get() and not self.step6_enabled.get():
            return "Step 7 を実行するには Step 6 を有効にしてください。"

        return None

    def prepare_intermediate_inputs(self):
        """Step 2 をスキップする場合の中間入力を初期化する。"""
        self.intermediate_contact_csv = None
        self.intermediate_label_csv = None
        self.intermediate_contacts_csv = None
        self.intermediate_contactgroups_csv = None
        self.final_labels_csv_list = []

        if not self.step2_enabled.get():
            existing_contact_csv = self.existing_input_path(self.contact_csv_path.get())
            existing_label_csv = self.existing_input_path(self.label_csv_path.get())
            if existing_contact_csv:
                self.intermediate_contact_csv = existing_contact_csv
            if existing_label_csv:
                self.intermediate_label_csv = existing_label_csv

    def confirm_settings(self):
        """設定内容を表示"""
        message = self.generate_confirmation_message()
        messagebox.showinfo("設定内容の確認", message, icon='info')

    def confirm_and_run(self):
        """実行確認ダイアログ表示"""
        # バリデーション
        if not self.export_csv_path.get():
            messagebox.showerror("エラー", "エクスポートCSVを選択してください")
            return

        if not self.target_email.get():
            messagebox.showerror("エラー", "ターゲットメールアドレスを入力してください")
            return

        # メールアドレス形式チェック
        email_pattern = r'^[^\s@]+@[^\s@]+\.[^\s@]+$'
        if not re.match(email_pattern, self.target_email.get()):
            messagebox.showerror("エラー", "メールアドレスの形式が正しくありません")
            return

        dependency_error = self.validate_step_dependencies()
        if dependency_error:
            messagebox.showerror("エラー", dependency_error)
            return

        # 確認メッセージ
        confirm_msg = self.generate_confirmation_message()
        confirm_msg += "\n\n処理を開始してよろしいですか？"

        result = messagebox.askyesno("実行確認", confirm_msg, icon='question')

        if result:
            self.run_all_steps()

    def run_all_steps(self):
        """全ステップを順次実行"""
        # ボタン無効化
        self.process_button.config(state=tk.DISABLED, text="実行中...")

        # 別スレッドで実行
        thread = threading.Thread(target=self.process_worker)
        thread.daemon = True
        thread.start()

    def process_worker(self):
        """処理ワーカースレッド"""
        try:
            self.prepare_intermediate_inputs()

            # 作業ディレクトリ作成（絶対パスで作成）
            run_id = datetime.now().strftime('%Y%m%d_%H%M%S')
            self.work_dir = self.base_dir / "runs" / run_id
            self.work_dir.mkdir(parents=True, exist_ok=True)

            self.log("=" * 60)
            self.log(f"処理開始: {run_id}")
            self.log(f"作業ディレクトリ: {self.work_dir}")
            self.log("=" * 60)
            self.log("")

            # Step 2: CSV抽出
            if self.step2_enabled.get():
                self.execute_step2()
                self.log("")

            # Step 3: 連絡先登録
            if self.step3_enabled.get():
                self.execute_step3()
                self.log("")

            # Step 4: ラベル作成
            if self.step4_enabled.get():
                self.execute_step4()
                self.log("")

            # Step 5: 登録データ取得
            if self.step5_enabled.get():
                self.execute_step5()
                self.log("")

            # Step 6: データ突合
            if self.step6_enabled.get():
                self.execute_step6()
                self.log("")

            # Step 7: ラベル登録
            if self.step7_enabled.get():
                self.execute_step7()
                self.log("")

            self.log("=" * 60)
            self.log("✅ すべての処理が完了しました")
            self.log("=" * 60)

            # 完了ダイアログ
            self.root.after(0, lambda: messagebox.showinfo(
                "完了",
                "すべての処理が正常に完了しました！"
            ))

        except Exception as e:
            self.log(f"❌ エラー: {str(e)}")
            import traceback
            self.log(traceback.format_exc())

            self.root.after(0, lambda: messagebox.showerror(
                "エラー",
                f"処理中にエラーが発生しました:\n{str(e)}"
            ))

        finally:
            # ボタン有効化
            self.root.after(0, lambda: self.process_button.config(
                state=tk.NORMAL,
                text="処理開始 ▶"
            ))

    def run_gam(self, gam_args, timeout=600):
        """PowerShellラッパー経由でGAMを安全に実行する。"""
        if not self.gam_wrapper_script.exists():
            raise Exception(f"GAMラッパースクリプトが見つかりません: {self.gam_wrapper_script}")

        cmd = [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", str(self.gam_wrapper_script),
            "-GamArgsJson", json.dumps(gam_args, ensure_ascii=False)
        ]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            timeout=timeout,
            cwd=str(self.base_dir)
        )

        output_parts = []
        if result.stdout and result.stdout.strip():
            output_parts.append(result.stdout.strip())
        if result.stderr and result.stderr.strip():
            output_parts.append(result.stderr.strip())

        return result.returncode, "\n".join(output_parts)

    def execute_step2(self):
        """Step 2: CSV抽出"""
        self.log("--- Step 2: CSV抽出 開始 ---")

        try:
            # 出力先決定
            if self.contact_csv_path.get():
                contact_csv = Path(self.contact_csv_path.get())
            else:
                contact_csv = self.work_dir / "連絡先.csv"

            if self.label_csv_path.get():
                label_csv = Path(self.label_csv_path.get())
            else:
                label_csv = self.work_dir / "ラベル.csv"

            self.log(f"  入力: {self.export_csv_path.get()}")
            self.log(f"  出力: {contact_csv}, {label_csv}")

            # Python実行
            cmd = [
                self.python_executable,
                "csvContactsFirst/convert_contacts.py",
                self.export_csv_path.get(),
                str(contact_csv),
                str(label_csv),
                "--log-level", "INFO"
            ]

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=300,  # 5分タイムアウト
                cwd=str(self.base_dir)  # カレントディレクトリを指定
            )

            if result.returncode == 0:
                self.log("  " + result.stdout.replace('\n', '\n  '))
                self.log("✅ Step 2: 完了")

                # 中間ファイルパスを保存
                self.intermediate_contact_csv = contact_csv
                self.intermediate_label_csv = label_csv
            else:
                self.log(f"❌ Step 2: 失敗")
                self.log(f"  標準出力: {result.stdout}")
                self.log(f"  標準エラー: {result.stderr}")
                error_msg = result.stderr if result.stderr else result.stdout
                raise Exception(f"Step 2 失敗: {error_msg}")

        except subprocess.TimeoutExpired:
            self.log("❌ Step 2: タイムアウト（5分以上）")
            raise Exception("Step 2がタイムアウトしました")
        except Exception as e:
            self.log(f"❌ Step 2: エラー - {str(e)}")
            raise

    def execute_step3(self):
        """Step 3: 連絡先登録（PowerShell呼び出し）"""
        self.log("--- Step 3: 連絡先登録 開始 ---")

        try:
            self.log(f"  連絡先CSV: {self.intermediate_contact_csv}")
            self.log(f"  ターゲットメール: {self.target_email.get()}")

            # PowerShellスクリプトを呼び出し
            cmd = [
                "powershell.exe",
                "-ExecutionPolicy", "Bypass",
                "-File", "registerContacts/run_register_cli.ps1",
                "-ContactCsvFile", str(self.intermediate_contact_csv),
                "-TargetUserEmail", self.target_email.get(),
                "-LogDir", self.log_dir.get()
            ]

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=1800,  # 30分タイムアウト
                cwd=str(self.base_dir)  # カレントディレクトリを指定
            )

            if result.returncode == 0:
                self.log("  " + result.stdout.replace('\n', '\n  '))
                self.log("✅ Step 3: 完了")
            else:
                self.log(f"❌ Step 3: 失敗")
                self.log(f"  標準出力: {result.stdout}")
                self.log(f"  標準エラー: {result.stderr}")
                error_msg = result.stderr if result.stderr else result.stdout
                raise Exception(f"Step 3 失敗: {error_msg}")

        except subprocess.TimeoutExpired:
            self.log("❌ Step 3: タイムアウト（30分以上）")
            raise Exception("Step 3がタイムアウトしました")
        except Exception as e:
            self.log(f"❌ Step 3: エラー - {str(e)}")
            raise

    def execute_step4(self):
        """Step 4: ラベル作成（重複除去対応）"""
        self.log("--- Step 4: ラベル作成 開始 ---")

        try:
            self.log(f"  ラベルCSV: {self.intermediate_label_csv}")
            self.log(f"  ターゲットメール: {self.target_email.get()}")

            # CSVを読み込み、すべてのラベル名を収集
            import csv
            unique_labels = set()

            with open(self.intermediate_label_csv, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # 各列からラベル名を収集
                    for column_name, label_name in row.items():
                        if not label_name or not label_name.strip():
                            continue

                        label_name = label_name.strip()

                        # "* myContacts"は除外（システムラベル）
                        if label_name == '* myContacts':
                            continue

                        # "* "プレフィックスを削除
                        if label_name.startswith('* '):
                            label_name = label_name[2:].strip()

                        # 空でないラベル名のみ追加
                        if label_name and label_name != 'myContacts':
                            unique_labels.add(label_name)

            if not unique_labels:
                self.log("  警告: ラベルが見つかりませんでした")
            else:
                self.log(f"  検出されたユニークなラベル: {len(unique_labels)}個")

                # ラベル名をソートして表示
                sorted_labels = sorted(unique_labels)
                for label in sorted_labels:
                    self.log(f"    - {label}")

                # 各ユニークなラベルに対してラベル作成コマンドを実行
                success_count = 0
                skip_count = 0

                for label_name in sorted_labels:
                    self.log(f"  作成中: {label_name}")
                    returncode, output = self.run_gam(
                        ["user", self.target_email.get(), "create", "contactgroup", "name", label_name]
                    )

                    result_lower = output.lower() if output else ""
                    if returncode != 0:
                        if 'already exists' in result_lower or '既に存在' in result_lower:
                            self.log(f"    スキップ（既存）: {label_name}")
                            skip_count += 1
                        else:
                            self.log(f"    ⚠️ エラー: {output or 'GAM実行失敗'}")
                    else:
                        if output:
                            self.log("    GAM出力: " + output.replace('\n', '\n    '))
                        self.log(f"    ✅ 作成完了: {label_name}")
                        success_count += 1

                self.log(f"  ラベル作成完了（成功: {success_count}個, スキップ: {skip_count}個）")

            self.log("✅ Step 4: 完了")

        except Exception as e:
            self.log(f"⚠️ Step 4: 一部失敗の可能性 - {str(e)}")
            # ラベル作成は一部失敗しても処理を継続
            self.log("  処理を継続します")

    def execute_step5(self):
        """Step 5: 登録データ取得（Gitbash使用）"""
        self.log("--- Step 5: 登録データ取得 開始 ---")

        contacts_output = self.work_dir / "contacts.csv"
        labels_output = self.work_dir / "contactgroups.csv"

        try:
            # 連絡先一覧取得
            self.log("  連絡先一覧を取得中...")
            returncode, contacts_csv = self.run_gam(
                ["user", self.target_email.get(), "print", "contacts"]
            )
            if returncode != 0:
                raise Exception(contacts_csv or "contacts の取得に失敗しました")

            with open(contacts_output, 'w', encoding='utf-8', newline='') as f:
                f.write(contacts_csv)

            self.log(f"  ✅ 連絡先一覧: {contacts_output}")

            # ラベル一覧取得
            self.log("  ラベル一覧を取得中...")
            returncode, labels_csv = self.run_gam(
                ["user", self.target_email.get(), "print", "contactgroups"]
            )
            if returncode != 0:
                raise Exception(labels_csv or "contactgroups の取得に失敗しました")

            with open(labels_output, 'w', encoding='utf-8', newline='') as f:
                f.write(labels_csv)

            self.log(f"  ✅ ラベル一覧: {labels_output}")
            self.log("✅ Step 5: 完了")

            self.intermediate_contacts_csv = contacts_output
            self.intermediate_contactgroups_csv = labels_output

        except Exception as e:
            self.log(f"❌ Step 5: 失敗 - {str(e)}")
            raise

    def execute_step6(self):
        """Step 6: データ突合（ラベル数ごとに分割されたCSV生成）"""
        self.log("--- Step 6: データ突合 開始 ---")

        try:
            output_csv_base = self.work_dir / "contacts_labels"

            self.log(f"  エクスポートCSV: {self.export_csv_path.get()}")
            self.log(f"  連絡先CSV: {self.intermediate_contact_csv}")
            self.log(f"  contacts.csv: {self.intermediate_contacts_csv}")
            self.log(f"  contactgroups.csv: {self.intermediate_contactgroups_csv}")
            self.log(f"  出力ベース: {output_csv_base}")

            cmd = [
                self.python_executable,
                "csvContactsSecond/main.py",
                "--export-data", self.export_csv_path.get(),
                "--registered-data", str(self.intermediate_contact_csv),
                "--label-data", str(self.intermediate_contactgroups_csv),
                "--contacts-data", str(self.intermediate_contacts_csv),
                "--target-email", self.target_email.get(),
                "--output", str(output_csv_base) + ".csv"  # ベース名を渡す
            ]

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=300,
                cwd=str(self.base_dir)  # 作業ディレクトリを明示的に設定
            )

            if result.returncode == 0:
                self.log("  " + result.stdout.replace('\n', '\n  '))

                # 生成されたCSVファイルを検出
                import glob
                pattern = str(self.work_dir / "contacts_labels_*.csv")
                generated_files = sorted(glob.glob(pattern))

                if generated_files:
                    self.final_labels_csv_list = [Path(f) for f in generated_files]
                    self.log(f"  生成されたCSVファイル: {len(self.final_labels_csv_list)}個")
                    for csv_file in self.final_labels_csv_list:
                        self.log(f"    - {csv_file.name}")
                else:
                    self.log("  ⚠️ CSVファイルが見つかりませんでした")

                self.log("✅ Step 6: 完了")
            else:
                self.log(f"❌ Step 6: 失敗")
                self.log(f"  標準出力: {result.stdout}")
                self.log(f"  標準エラー: {result.stderr}")
                error_msg = result.stderr if result.stderr else result.stdout
                raise Exception(f"Step 6 失敗: {error_msg}")

        except subprocess.TimeoutExpired:
            self.log("❌ Step 6: タイムアウト（5分以上）")
            raise Exception("Step 6がタイムアウトしました")
        except Exception as e:
            self.log(f"❌ Step 6: エラー - {str(e)}")
            raise

    def execute_step7(self):
        """Step 7: ラベル登録（複数CSVファイル対応）"""
        self.log("--- Step 7: ラベル登録 開始 ---")

        try:
            if not self.final_labels_csv_list:
                self.log("  ⚠️ ラベル登録用CSVファイルが見つかりません")
                return

            self.log(f"  処理するCSVファイル数: {len(self.final_labels_csv_list)}")

            success_count = 0
            error_count = 0

            # 各CSVファイルに対してGAMコマンドを実行
            for csv_file in self.final_labels_csv_list:
                self.log("")
                self.log(f"  処理中: {csv_file.name}")

                try:
                    # CSVヘッダー行を読み取り、ラベル列を自動検出
                    import csv
                    with open(csv_file, 'r', encoding='utf-8') as f:
                        reader = csv.reader(f)
                        headers = next(reader)
                        # データ行数を数える
                        row_count = sum(1 for _ in reader)

                    # すべてのラベル列を検出（PrimaryLabel含む、TargetEmail/ContactIDは除外）
                    label_columns = [
                        col for col in headers
                        if col.endswith('Label')
                        and col not in ['TargetEmail', 'ContactID']
                    ]

                    if not label_columns:
                        self.log(f"    ⚠️ ラベル列が見つかりませんでした: {csv_file.name}")
                        error_count += 1
                        continue

                    self.log(f"    対象レコード数: {row_count}件")
                    self.log(f"    ラベル列: {', '.join(label_columns)}")

                    gam_args = [
                        "csv", str(csv_file),
                        "gam", "user", "~TargetEmail",
                        "update", "contacts", "~ContactID"
                    ]
                    for column_name in label_columns:
                        gam_args.extend(["contactgroup", f"~{column_name}"])

                    self.log(f"    実行引数: {' '.join(gam_args)}")

                    returncode, output = self.run_gam(gam_args)
                    if returncode != 0:
                        raise Exception(output or "GAM実行失敗")

                    if output:
                        self.log("    GAM出力: " + output.replace('\n', '\n    '))

                    self.log(f"    ✅ 完了: {csv_file.name}")
                    success_count += 1

                except Exception as e:
                    self.log(f"    ❌ エラー: {csv_file.name} - {str(e)}")
                    error_count += 1

            self.log("")
            self.log(f"  処理結果: 成功 {success_count}件, エラー {error_count}件")
            self.log("✅ Step 7: 完了")

        except Exception as e:
            self.log(f"⚠️ Step 7: エラー - {str(e)}")
            self.log("  処理を継続します")


def main():
    """メイン処理"""
    root = tk.Tk()
    app = TransferContactsMasterGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
