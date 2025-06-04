import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import subprocess
import csv
import os
from datetime import datetime
import win32evtlog
import win32evtlogutil
import win32con
import threading
import sys

class ModernILCollector:
    WINDOW_TITLE_SUFFIX = "ILCollector - イベントログ収集ツール"

    def __init__(self):
        # メインウィンドウの作成
        self.root = tk.Tk()
        self.root.title(self.WINDOW_TITLE_SUFFIX)
        self.root.geometry("800x600")
        self.root.minsize(400, 300)  # 最小サイズを設定
        self.root.resizable(True, True)
        
        # モダンな配色テーマ
        self.colors = {
            'bg_primary': '#1a1a1a',      # ダークグレー
            'bg_secondary': '#2d2d2d',    # ライトダークグレー
            'bg_card': '#3a3a3a',         # カード背景
            'accent_blue': '#0078d4',     # モダンブルー
            'accent_green': '#107c10',    # モダングリーン
            'accent_orange': '#ff8c00',   # モダンオレンジ
            'accent_yellow': '#ffb900',   # モダンイエロー
            'accent_gray': '#2d2d2d',     # グレー
            'text_primary': '#ffffff',    # 白文字
            'text_secondary': '#cccccc',  # グレー文字
            'border': '#4a4a4a'           # ボーダー色
        }

        # フォント
        self.font_family = 'メイリオ'
        
        # ウィンドウの背景色を設定
        self.root.configure(bg=self.colors['bg_primary'])
        
        # 処理中メッセージウィンドウ用
        self.progress_window = None
        
        # 出力フォルダの設定
        current_dir = os.getcwd()
        timestamp = datetime.now().strftime("%Y%m%d-%H%M")
        self.output_folder = os.path.join(current_dir, timestamp)
        
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)
        
        self.setup_styles()
        self.create_modern_widgets()
    
    def setup_styles(self):
        """モダンなスタイルを設定"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # カスタムボタンスタイル
        self.style.configure('Modern.TButton', 
                            borderwidth=0,
                            relief='flat',
                            padding=(40, 25),  # 左右のパディングを20→30、上下を15→20に増加
                            font=(self.font_family, 12, 'bold'))  # フォントサイズを11→12に増加
        
        # プライマリボタン（青）
        self.style.configure('Primary.TButton', 
                            background=self.colors['accent_blue'],
                            foreground='white',
                            focuscolor='none',
                            padding=(40, 25),  # パディング設定を追加
                            font=(self.font_family, 12, 'bold'))  # フォント設定を追加)
        
        # 成功ボタン（緑）
        self.style.configure('Success.TButton', 
                            background=self.colors['accent_green'],
                            foreground='white',
                            focuscolor='none',
                            padding=(40, 25),  # パディング設定を追加
                            font=(self.font_family, 12, 'bold'))  # フォント設定を追加)
        
        # 警告ボタン（オレンジ）
        self.style.configure('Warning.TButton', 
                            background=self.colors['accent_orange'],
                            foreground='white',
                            focuscolor='none',
                            padding=(40, 25),  # パディング設定を追加
                            font=(self.font_family, 12, 'bold'))  # フォント設定を追加)
        
        # フォルダボタン（黄）
        self.style.configure('Folder.TButton', 
                            background=self.colors['accent_yellow'],
                            foreground='black',
                            focuscolor='none',
                            padding=(40, 25),  # パディング設定を追加
                            font=(self.font_family, 12, 'bold'))  # フォント設定を追加)
    
    def create_modern_widgets(self):
        """モダンなUIコンポーネントを作成"""
        # スクロール可能なキャンバスとスクロールバーを作成
        canvas = tk.Canvas(self.root, bg=self.colors['bg_primary'], highlightthickness=0)
        scrollbar = tk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['bg_primary'])
        
        # スクロール可能フレームの設定
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # キャンバスとスクロールバーの配置
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # メインコンテナをスクロール可能フレーム内に作成
        main_container = tk.Frame(scrollable_frame, bg=self.colors['bg_primary'])
        main_container.pack(padx=30, pady=30)  # fill="both", expand=Trueを削除

        # 中央配置用の外側フレームを追加
        center_frame = tk.Frame(scrollable_frame, bg=self.colors['bg_primary'])
        center_frame.pack(fill="both", expand=True)

        main_container = tk.Frame(center_frame, bg=self.colors['bg_primary'])
        main_container.pack(padx=30, pady=30, anchor="center")
        
        # ヘッダーセクション
        self.create_header(main_container)
        
        # カードコンテナ
        cards_container = tk.Frame(main_container, bg=self.colors['bg_primary'])
        cards_container.pack(fill="both", expand=True, pady=(30, 0))
        
        # 機能カード
        self.create_feature_cards(cards_container)
        
        # フッターセクション
        self.create_footer(main_container)
        
        # マウスホイールスクロールの設定
        self.bind_mousewheel(canvas)

    # マウスホイールスクロール用のメソッドを追加
    def bind_mousewheel(self, canvas):
        """マウスホイールでのスクロールを有効化"""
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_to_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        def _unbind_from_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
        
        canvas.bind('<Enter>', _bind_to_mousewheel)
        canvas.bind('<Leave>', _unbind_from_mousewheel)
    
    def create_header(self, parent):
        """ヘッダーセクションを作成"""
        header_frame = tk.Frame(parent, bg=self.colors['bg_primary'])
        header_frame.pack(fill="x", pady=(0, 20))
        
        # タイトル
        title_label = tk.Label(
            header_frame,
            text="🔍 ILCollector",
            font=(self.font_family, 28, 'bold'),
            fg=self.colors['text_primary'],
            bg=self.colors['bg_primary']
        )
        title_label.pack()
        
        # サブタイトル
        subtitle_label = tk.Label(
            header_frame,
            text="Windows Server イベントログ & システム情報収集ツール",
            font=(self.font_family, 12),
            fg=self.colors['text_secondary'],
            bg=self.colors['bg_primary']
        )
        subtitle_label.pack(pady=(5, 0))

        # 区切り線
        separator = tk.Frame(header_frame, height=1, bg=self.colors['accent_gray'])
        separator.pack(fill="x", pady=(20, 0))

        # 注意文
        subtitle_label = tk.Label(
            header_frame,
            text="※本ツールは弊社による障害対応を支援・補助するものですが、復旧を保証するものではありません。",
            font=(self.font_family, 10),
            fg=self.colors['text_secondary'],
            bg=self.colors['bg_primary']
        )
        subtitle_label.pack(pady=(5, 0))
        
        # 区切り線
        separator = tk.Frame(header_frame, height=2, bg=self.colors['accent_blue'])
        separator.pack(fill="x", pady=(20, 0))
    
    def create_feature_cards(self, parent):
        """機能カードを作成"""
        # カードのグリッド配置用フレーム
        cards_frame = tk.Frame(parent, bg=self.colors['bg_primary'])
        cards_frame.pack(expand=True, fill="both")
        
        # カード1: イベントログ出力
        card1 = self.create_card(
            cards_frame,
            "📊 イベントログ出力",
            "System・Applicationログを\nCSVファイルとして出力",
            "Primary.TButton",
            self.export_eventlogs,
            row=0, col=0, colspan=1
        )

        # カード2: システム情報出力
        card2 = self.create_card(
            cards_frame,
            "💻 システム情報出力",
            "CPU・メモリ・OS情報などの\n詳細情報をテキストで出力",
            "Success.TButton",
            self.export_msinfo,
            row=1, col=0, colspan=1
        )

        # カード3: 一括取得
        card3 = self.create_card(
            cards_frame,
            "🚀 すべて一括取得",
            "上記の処理をまとめて実行\n（推奨オプション）",
            "Warning.TButton",
            self.export_all,
            row=2, col=0, colspan=1, large=True
        )

        # グリッドの重み設定を1列に変更
        cards_frame.grid_columnconfigure(0, weight=1)
        cards_frame.grid_rowconfigure(0, weight=1)
        cards_frame.grid_rowconfigure(1, weight=1)
        cards_frame.grid_rowconfigure(2, weight=1)
    
    def create_card(self, parent, title, description, button_style, command, 
                   row, col, colspan=1, large=False):
        """モダンなカードコンポーネントを作成"""
        # カードフレーム
        card_frame = tk.Frame(
            parent, 
            bg=self.colors['bg_card'],
            relief='flat',
            bd=1
        )
        
        # グリッド配置
        padx = 10 if colspan == 1 else 0
        pady = 10
        card_frame.grid(row=row, column=col, columnspan=colspan, 
                       sticky="nsew", padx=padx, pady=pady)
        
        # カード内のコンテンツ
        content_frame = tk.Frame(card_frame, bg=self.colors['bg_card'])
        content_frame.pack(expand=True, fill="both", padx=25, pady=25)
        
        # タイトル
        title_label = tk.Label(
            content_frame,
            text=title,
            font=(self.font_family, 16 if large else 14, 'bold'),
            fg=self.colors['text_primary'],
            bg=self.colors['bg_card']
        )
        title_label.pack(pady=(0, 10))
        
        # 説明文
        desc_label = tk.Label(
            content_frame,
            text=description,
            font=(self.font_family, 10),
            fg=self.colors['text_secondary'],
            bg=self.colors['bg_card'],
            justify="center"
        )
        desc_label.pack(pady=(0, 20))
        
        # 実行ボタン
        button = ttk.Button(
            content_frame,
            text="実行",
            style=button_style,
            command=command
        )
        button.pack()
        
        return card_frame
    
    def create_footer(self, parent):
        """フッターセクションを作成"""
        footer_frame = tk.Frame(parent, bg=self.colors['bg_primary'])
        footer_frame.pack(fill="x", pady=(30, 0))
        
        # ボタンコンテナ
        button_container = tk.Frame(footer_frame, bg=self.colors['bg_primary'])
        button_container.pack()
        
        # フォルダを開くボタン
        folder_btn = ttk.Button(
            button_container,
            text="📁 出力フォルダを開く",
            style="Folder.TButton",
            command=self.open_output_folder
        )
        folder_btn.pack(side="left", padx=(0, 15))
        
        # 終了ボタン
        exit_btn = ttk.Button(
            button_container,
            text="❌ 終了",
            command=self.root.quit
        )
        exit_btn.pack(side="left")
        
        # 出力パス表示
        self.folder_label = tk.Label(
            footer_frame,
            text=f"出力先: {self.output_folder}",
            font=(self.font_family, 9),
            fg=self.colors['text_secondary'],
            bg=self.colors['bg_primary'],
            wraplength=740,
            justify="center"
        )
        self.folder_label.pack(pady=(15, 0))

        # コピーライト
        self.folder_label = tk.Label(
            footer_frame,
            text=f"© 2025 Reishin Sakuma",
            font=(self.font_family, 9),
            fg=self.colors['text_secondary'],
            bg=self.colors['bg_primary'],
            wraplength=740,
            justify="center"
        )
        self.folder_label.pack(pady=(15, 0))
    
    def open_output_folder(self):
        """出力フォルダをエクスプローラーで開く"""
        try:
            os.startfile(self.output_folder)
        except Exception as e:
            self.show_modern_error("エラー", f"フォルダを開けませんでした:\n{str(e)}")
    
    def show_modern_completion(self, title, message, files_info):
        """モダンな完了ダイアログを表示"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"{title} - {self.WINDOW_TITLE_SUFFIX}")
        # dialog.geometry("500x350")  # ← この行を削除
        dialog.resizable(False, False)
        dialog.configure(bg=self.colors['bg_primary'])
        dialog.transient(self.root)
        dialog.grab_set()
        
        # メインコンテンツ
        content_frame = tk.Frame(dialog, bg=self.colors['bg_card'])
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 成功アイコンとタイトル
        header_frame = tk.Frame(content_frame, bg=self.colors['bg_card'])
        header_frame.pack(fill="x", pady=(20, 15))
        
        success_label = tk.Label(
            header_frame,
            text="✅",
            font=(self.font_family, 32),
            bg=self.colors['bg_card']
        )
        success_label.pack()
        
        title_label = tk.Label(
            header_frame,
            text=title,
            font=(self.font_family, 16, 'bold'),
            fg=self.colors['text_primary'],
            bg=self.colors['bg_card']
        )
        title_label.pack(pady=(10, 0))
        
        # メッセージ
        msg_label = tk.Label(
            content_frame,
            text=message,
            font=(self.font_family, 11),
            fg=self.colors['text_secondary'],
            bg=self.colors['bg_card']
        )
        msg_label.pack(pady=(0, 15))
        
        # ファイル情報
        files_label = tk.Label(
            content_frame,
            text=files_info,
            font=(self.font_family, 9),
            fg=self.colors['text_secondary'],
            bg=self.colors['bg_card'],
            justify="left",
            wraplength=460
        )
        files_label.pack(pady=(5, 10))
        
        # ボタンフレーム
        button_frame = tk.Frame(content_frame, bg=self.colors['bg_card'])
        button_frame.pack(pady=(10, 20))
        
        # フォルダを開くボタン
        open_btn = ttk.Button(
            button_frame,
            text="📁 フォルダを開く",
            style="Folder.TButton",
            command=lambda: [self.open_output_folder()]
        )
        open_btn.pack(side="left", padx=(0, 10))
        
        # OKボタン
        ok_btn = ttk.Button(
            button_frame,
            text="OK",
            style="Primary.TButton",
            command=dialog.destroy
        )
        ok_btn.pack(side="left")

        # 内容に合わせてウィンドウサイズを自動調整
        dialog.update_idletasks()
        dialog.geometry("")  # サイズ自動調整

        # 中央配置
        self.center_window(dialog)
    
    def show_modern_progress(self, message):
        """モダンな処理中ダイアログを表示"""
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title(f"処理中 - {self.WINDOW_TITLE_SUFFIX}")
        self.progress_window.geometry("400x200")  # 幅350→400px、高さ150→200pxに変更
        self.progress_window.resizable(False, False)
        self.progress_window.configure(bg=self.colors['bg_primary'])
        self.progress_window.transient(self.root)
        self.progress_window.grab_set()
        
        # 中央配置
        self.center_window(self.progress_window)
        
        # コンテンツフレーム
        content_frame = tk.Frame(self.progress_window, bg=self.colors['bg_card'])
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # アニメーション風のアイコン
        icon_label = tk.Label(
            content_frame,
            text="⚙️",
            font=(self.font_family, 24),
            bg=self.colors['bg_card']
        )
        icon_label.pack(pady=(15, 15))  # 上下の余白を20,10→15,15に調整
        
        # メッセージ
        msg_label = tk.Label(
            content_frame,
            text=message,
            font=(self.font_family, 11),
            fg=self.colors['text_primary'],
            bg=self.colors['bg_card'],
            wraplength=350,  # 文字折り返し幅を300→350に拡大
            justify="center"
        )
        msg_label.pack(pady=(0, 15))  # 下の余白を20→15に調整    
    def show_modern_error(self, title, message):
        """モダンなエラーダイアログを表示"""
        messagebox.showerror(f"{title} - {self.WINDOW_TITLE_SUFFIX}", message)
    
    def center_window(self, window):
        """ウィンドウを親ウィンドウの中央に配置"""
        window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (window.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (window.winfo_height() // 2)
        window.geometry(f"+{x}+{y}")
    
    def hide_progress(self):
        """処理中ダイアログを閉じる"""
        if self.progress_window:
            self.progress_window.destroy()
            self.progress_window = None
    
    # 以下、元のメソッドをモダンUI対応に修正
    def export_eventlogs(self):
        """イベントログをCSVに出力する（メインスレッド）"""
        thread = threading.Thread(target=self._export_eventlogs_thread)
        thread.daemon = True
        thread.start()
    
    def _export_eventlogs_thread(self):
        """イベントログをCSVに出力する（バックグラウンド処理）"""
        try:
            self.root.after(0, lambda: self.show_modern_progress("イベントログを収集しています...\n少々お待ちください。"))
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            system_file = os.path.join(self.output_folder, f"System_EventLog_{timestamp}.csv")
            self.get_eventlog("System", system_file)
            
            app_file = os.path.join(self.output_folder, f"Application_EventLog_{timestamp}.csv")
            self.get_eventlog("Application", app_file)
            
            self.root.after(0, self.hide_progress)
            
            files_info = f"出力ファイル:\n• {os.path.basename(system_file)}\n• {os.path.basename(app_file)}\n\n出力先:\n{self.output_folder}"
            self.root.after(0, lambda: self.show_modern_completion(
                "処理完了",
                "イベントログの出力が完了しました！",
                files_info
            ))
            
        except Exception as e:
            error_msg = f"イベントログの出力中にエラーが発生しました:\n{str(e)}"  # エラーメッセージを変数に保存
            self.root.after(0, self.hide_progress)
            self.root.after(0, lambda msg=error_msg: self.show_modern_error("エラー", msg))  # デフォルト引数で値を固定
    
    def export_all(self):
        """すべてのログ・情報を一括取得する（メインスレッド）"""
        thread = threading.Thread(target=self._export_all_thread)
        thread.daemon = True
        thread.start()
    
    def _export_all_thread(self):
        """すべてのログ・情報を一括取得する（バックグラウンド処理）"""
        try:
            self.root.after(0, lambda: self.show_modern_progress("すべてのログ・情報を収集しています...\n少々お待ちください。"))
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            files_created = []
            
            # イベントログの出力
            system_file = os.path.join(self.output_folder, f"System_EventLog_{timestamp}.csv")
            self.get_eventlog("System", system_file)
            files_created.append(os.path.basename(system_file))
            
            app_file = os.path.join(self.output_folder, f"Application_EventLog_{timestamp}.csv")
            self.get_eventlog("Application", app_file)
            files_created.append(os.path.basename(app_file))
            
            # システム情報の出力
            output_file = os.path.join(self.output_folder, f"SystemInfo_{timestamp}.txt")
            
            cmd = f'msinfo32 /report "{output_file}"'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            
            if result.returncode != 0:
                result = subprocess.run(
                    'systeminfo', 
                    shell=True, 
                    capture_output=True, 
                    text=True, 
                    encoding='shift_jis'
                )
                
                if result.returncode == 0:
                    with open(output_file, 'w', encoding='utf-8') as f:
                        f.write("=== システム情報 ===\n")
                        f.write(result.stdout)
            
            files_created.append(os.path.basename(output_file))
            
            self.root.after(0, self.hide_progress)
            
            files_info = "出力ファイル:\n" + "\n".join([f"• {f}" for f in files_created]) + f"\n\n出力先:\n{self.output_folder}"
            self.root.after(0, lambda: self.show_modern_completion(
                "処理完了",
                "すべてのログ・情報の出力が完了しました！",
                files_info
            ))
            
        except Exception as e:
            error_msg = f"ログ・情報の出力中にエラーが発生しました:\n{str(e)}"  # エラーメッセージを変数に保存
            self.root.after(0, self.hide_progress)
            self.root.after(0, lambda msg=error_msg: self.show_modern_error("エラー", msg))  # デフォルト引数で値を固定
    
    def export_msinfo(self):
        """msinfo32の情報をファイルに出力する（メインスレッド）"""
        thread = threading.Thread(target=self._export_msinfo_thread)
        thread.daemon = True
        thread.start()
    
    def _export_msinfo_thread(self):
        """msinfo32の情報をファイルに出力する（バックグラウンド処理）"""
        try:
            self.root.after(0, lambda: self.show_modern_progress("システム情報を収集しています...\n少々お待ちください。"))
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(self.output_folder, f"SystemInfo_{timestamp}.txt")
            
            cmd = f'msinfo32 /report "{output_file}"'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            
            if result.returncode == 0:
                self.root.after(0, self.hide_progress)
                files_info = f"出力ファイル:\n• {os.path.basename(output_file)}\n\n出力先:\n{self.output_folder}"
                self.root.after(0, lambda: self.show_modern_completion(
                    "処理完了",
                    "システム情報の出力が完了しました！",
                    files_info
                ))
            else:
                self.export_systeminfo_alternative(output_file)
                
        except Exception as e:
            error_msg = f"システム情報の出力中にエラーが発生しました:\n{str(e)}"  # エラーメッセージを変数に保存
            self.root.after(0, self.hide_progress)
            self.root.after(0, lambda msg=error_msg: self.show_modern_error("エラー", msg))  # デフォルト引数で値を固定
    
    def export_systeminfo_alternative(self, output_file):
        """代替方法でシステム情報を出力する"""
        try:
            result = subprocess.run(
                'systeminfo', 
                shell=True, 
                capture_output=True, 
                text=True, 
                encoding='shift_jis'
            )
            
            if result.returncode == 0:
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write("=== システム情報 ===\n")
                    f.write(result.stdout)
                
                self.root.after(0, self.hide_progress)
                files_info = f"出力ファイル:\n• {os.path.basename(output_file)}\n\n出力先:\n{self.output_folder}"
                self.root.after(0, lambda: self.show_modern_completion(
                    "処理完了",
                    "システム情報の出力が完了しました！",
                    files_info
                ))
            else:
                raise Exception("systeminfoコマンドも失敗しました")
                
        except Exception as e:
            error_msg = f"システム情報の取得に失敗しました:\n{str(e)}"  # エラーメッセージを変数に保存
            self.root.after(0, self.hide_progress)
            self.root.after(0, lambda msg=error_msg: self.show_modern_error("エラー", msg))  # デフォルト引数で値を固定
    
    def get_eventlog(self, log_name, output_file):
        """指定されたイベントログを取得してCSVに保存する"""
        hand = win32evtlog.OpenEventLog(None, log_name)
        
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            
            writer.writerow([
                '日時', 'イベントID', 'レベル', 'ソース', 'メッセージ'
            ])
            
            events = win32evtlog.ReadEventLog(
                hand, 
                win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ,
                0
            )
            
            count = 0
            while events and count < 1000:
                for event in events:
                    time_generated = event.TimeGenerated.Format()
                    
                    if event.EventType == win32con.EVENTLOG_ERROR_TYPE:
                        level = "エラー"
                    elif event.EventType == win32con.EVENTLOG_WARNING_TYPE:
                        level = "警告"
                    elif event.EventType == win32con.EVENTLOG_INFORMATION_TYPE:
                        level = "情報"
                    else:
                        level = "その他"
                    
                    try:
                        message = win32evtlogutil.SafeFormatMessage(event, log_name)
                        if message is None:
                            message = "メッセージを取得できませんでした"
                    except:
                        message = "メッセージを取得できませんでした"
                    
                    writer.writerow([
                        time_generated,
                        event.EventID & 0xFFFF,
                        level,
                        event.SourceName,
                        message.replace('\n', ' ').replace('\r', '')
                    ])
                    
                    count += 1
                    if count >= 1000:
                        break
                
                if count < 1000:
                    events = win32evtlog.ReadEventLog(
                        hand, 
                        win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ,
                        0
                    )
                else:
                    break
        
        win32evtlog.CloseEventLog(hand)
    
    def run(self):
        """アプリケーションを実行する"""
        self.root.mainloop()

# メイン実行部分
if __name__ == "__main__":
    # 管理者権限の確認
    import ctypes

    # グローバル変数として宣言
    should_start_app = True

    def show_admin_dialog():
        # 関数内でグローバル変数を使う
        global should_start_app
        root = tk.Tk()
        root.withdraw()  # メインウィンドウ非表示

        def on_continue():
            global should_start_app
            should_start_app = True
            root.quit()
            root.destroy()

        def on_exit():
            global should_start_app
            should_start_app = False
            root.quit()
            root.destroy()

        dialog = tk.Toplevel()
        dialog.title(f"権限不足 - ILCollector - イベントログ収集ツール")
        dialog.geometry("400x180")
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.configure(bg="#2d2d2d")

        label = tk.Label(
            dialog,
            text="このプログラムは管理者権限で実行することを推奨します。\n一部の機能が正常に動作しない可能性があります。",
            bg="#2d2d2d",
            fg="#ffffff",
            font=("メイリオ", 11),
            wraplength=360,
            justify="center"
        )
        label.pack(pady=(30, 20))

        btn_frame = tk.Frame(dialog, bg="#2d2d2d")
        btn_frame.pack(pady=(0, 20))

        continue_btn = ttk.Button(btn_frame, text="続行", command=on_continue)
        continue_btn.pack(side="left", padx=15)
        exit_btn = ttk.Button(btn_frame, text="終了", command=on_exit)
        exit_btn.pack(side="left", padx=15)

        root.mainloop()

    try:
        if not ctypes.windll.shell32.IsUserAnAdmin():
            show_admin_dialog()
    except Exception:
        pass

    # アプリケーションの起動
    if should_start_app:
        app = ModernILCollector()
        app.run()