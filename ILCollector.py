import tkinter as tk
from tkinter import messagebox, filedialog
import subprocess
import csv
import os
from datetime import datetime
import win32evtlog
import win32evtlogutil
import win32con
import threading

class ILCollector:
    def __init__(self):
        # メインウィンドウの作成
        self.root = tk.Tk()
        self.root.title("ILCollector - イベントログ収集ツール")
        self.root.geometry("700x500")  # 幅と高さを両方とも広げる
        self.root.resizable(True, True)  # 縦横両方向のリサイズを許可
        
        # 処理中メッセージウィンドウ用
        self.progress_window = None
        
        # 出力フォルダの設定（カレントディレクトリに作成）
        current_dir = os.getcwd()
        timestamp = datetime.now().strftime("%Y%m%d-%H%M")
        self.output_folder = os.path.join(current_dir, timestamp)
        
        # 出力フォルダが存在しない場合は作成
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)
        
        self.create_widgets()
        self.adjust_window_size()
    
    def create_widgets(self):
        """画面の部品を作成する"""
        # タイトルラベル
        title_label = tk.Label(
            self.root, 
            text="ILCollector", 
            font=("Arial", 16, "bold"),
            fg="blue"
        )
        title_label.pack(pady=20)
        
        # 説明ラベル
        desc_label = tk.Label(
            self.root,
            text="WindowsServerのイベントログとシステム情報を収集します",
            font=("Arial", 10)
        )
        desc_label.pack(pady=10)
        
        # イベントログ出力ボタンと説明
        eventlog_frame = tk.Frame(self.root)
        eventlog_frame.pack(pady=10, padx=20, fill="x")
        
        eventlog_btn = tk.Button(
            eventlog_frame,
            text="イベントログをCSV出力",
            font=("Arial", 12),
            bg="lightblue",
            width=25,
            height=2,
            command=self.export_eventlogs
        )
        eventlog_btn.pack(side=tk.LEFT)
        
        eventlog_desc = tk.Label(
            eventlog_frame,
            text="WindowsのSystemログとApplicationログを\nCSVファイルとして出力します",
            font=("Arial", 9),
            fg="gray",
            justify="left"
        )
        eventlog_desc.pack(side=tk.LEFT, padx=(15, 0))
        
        # msinfo32出力ボタンと説明
        msinfo_frame = tk.Frame(self.root)
        msinfo_frame.pack(pady=10, padx=20, fill="x")
        
        msinfo_btn = tk.Button(
            msinfo_frame,
            text="システム情報を出力",
            font=("Arial", 12),
            bg="lightgreen",
            width=25,
            height=2,
            command=self.export_msinfo
        )
        msinfo_btn.pack(side=tk.LEFT)
        
        msinfo_desc = tk.Label(
            msinfo_frame,
            text="CPU、メモリ、OS情報などの\nシステム詳細情報をテキストファイルで出力します",
            font=("Arial", 9),
            fg="gray",
            justify="left"
        )
        msinfo_desc.pack(side=tk.LEFT, padx=(15, 0))
        
        # 一括取得ボタンと説明
        batch_frame = tk.Frame(self.root)
        batch_frame.pack(pady=10, padx=20, fill="x")
        
        batch_btn = tk.Button(
            batch_frame,
            text="すべてのログ・情報を一括取得",
            font=("Arial", 12, "bold"),
            bg="orange",
            width=25,
            height=2,
            command=self.export_all
        )
        batch_btn.pack(side=tk.LEFT)
        
        batch_desc = tk.Label(
            batch_frame,
            text="上記の2つの処理を\nまとめて実行します（推奨）",
            font=("Arial", 9),
            fg="gray",
            justify="left"
        )
        batch_desc.pack(side=tk.LEFT, padx=(15, 0))
        
        # 出力フォルダを開くボタン（新規追加）
        folder_btn = tk.Button(
            self.root,
            text="📁 出力フォルダを開く",
            font=("Arial", 12),
            bg="lightyellow",
            width=25,
            height=2,
            command=self.open_output_folder
        )
        folder_btn.pack(pady=15)
        
        # 出力フォルダ表示（改善）
        self.folder_label = tk.Label(
            self.root,
            text=f"出力先: {self.output_folder}",
            font=("Arial", 9),
            fg="gray",
            wraplength=680,  # 文字列の折り返し幅を設定
            justify="center"
        )
        self.folder_label.pack(pady=10)
        
        # 終了ボタン
        exit_btn = tk.Button(
            self.root,
            text="終了",
            font=("Arial", 10),
            command=self.root.quit,
            bg="lightcoral"
        )
        exit_btn.pack(pady=10)
    
    def adjust_window_size(self):
        """パスの長さに応じてウィンドウサイズを調整"""
        # 出力パスの文字数を測定
        path_length = len(self.output_folder)
        
        # 基本サイズ
        base_width = 700
        base_height = 500
        
        # パスの長さに応じて幅を調整（1文字あたり約6ピクセル）
        if path_length > 90:
            additional_width = (path_length - 90) * 6
            new_width = min(base_width + additional_width, 1000)  # 最大1000ピクセル
            
            # 高さも少し調整（説明文が増えたため）
            new_height = min(base_height + 50, 600)  # 最大600ピクセル
            
            self.root.geometry(f"{new_width}x{new_height}")
            
            # ラベルの折り返し幅も調整
            self.folder_label.config(wraplength=new_width - 20)
    
    def open_output_folder(self):
        """出力フォルダをエクスプローラーで開く"""
        try:
            os.startfile(self.output_folder)
        except Exception as e:
            messagebox.showerror("エラー", f"フォルダを開けませんでした:\n{str(e)}")
    
    def show_completion_message(self, title, message, files_info):
        """完了メッセージをフォルダ開きボタン付きで表示"""
        # カスタムダイアログの作成
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("450x200")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 親ウィンドウの中央に配置
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # メッセージラベル
        msg_label = tk.Label(
            dialog,
            text=message + "\n\n" + files_info,
            font=("Arial", 10),
            wraplength=400,
            justify="left"
        )
        msg_label.pack(pady=15)
        
        # ボタンフレーム
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)
        
        # フォルダを開くボタン
        open_btn = tk.Button(
            button_frame,
            text="📁 出力フォルダを開く",
            font=("Arial", 10),
            bg="lightblue",
            command=lambda: [self.open_output_folder(), dialog.destroy()]
        )
        open_btn.pack(side=tk.LEFT, padx=10)
        
        # OKボタン
        ok_btn = tk.Button(
            button_frame,
            text="OK",
            font=("Arial", 10),
            command=dialog.destroy
        )
        ok_btn.pack(side=tk.LEFT, padx=10)
    
    def show_progress_message(self, message):
        """処理中メッセージを表示する"""
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("処理中")
        self.progress_window.geometry("300x100")
        self.progress_window.resizable(False, False)
        
        # 親ウィンドウの中央に表示
        self.progress_window.transient(self.root)
        self.progress_window.grab_set()
        
        # メッセージラベル
        label = tk.Label(
            self.progress_window,
            text=message,
            font=("Arial", 10),
            wraplength=250
        )
        label.pack(expand=True)
        
        # 親ウィンドウの中央に配置
        self.progress_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (self.progress_window.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (self.progress_window.winfo_height() // 2)
        self.progress_window.geometry(f"+{x}+{y}")
    
    def hide_progress_message(self):
        """処理中メッセージを閉じる"""
        if self.progress_window:
            self.progress_window.destroy()
            self.progress_window = None
    
    def export_eventlogs(self):
        """イベントログをCSVに出力する（メインスレッド）"""
        # 別スレッドで実行して画面をブロックしないようにする
        thread = threading.Thread(target=self._export_eventlogs_thread)
        thread.daemon = True
        thread.start()
    
    def _export_eventlogs_thread(self):
        """イベントログをCSVに出力する（バックグラウンド処理）"""
        try:
            # 処理中メッセージを表示
            self.root.after(0, lambda: self.show_progress_message("イベントログを収集しています...\n少々お待ちください。"))
            
            # 現在の日時を取得してファイル名に使用
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Systemログの出力
            system_file = os.path.join(self.output_folder, f"System_EventLog_{timestamp}.csv")
            self.get_eventlog("System", system_file)
            
            # Applicationログの出力
            app_file = os.path.join(self.output_folder, f"Application_EventLog_{timestamp}.csv")
            self.get_eventlog("Application", app_file)
            
            # 処理中メッセージを閉じる
            self.root.after(0, self.hide_progress_message)
            
            # 完了メッセージを表示（改善版）
            files_info = f"出力ファイル:\n- {os.path.basename(system_file)}\n- {os.path.basename(app_file)}\n\n出力先:\n{self.output_folder}"
            self.root.after(0, lambda: self.show_completion_message(
                "完了",
                "イベントログの出力が完了しました！",
                files_info
            ))
            
        except Exception as e:
            # 処理中メッセージを閉じる
            self.root.after(0, self.hide_progress_message)
            # エラーメッセージを表示
            self.root.after(0, lambda: messagebox.showerror("エラー", f"イベントログの出力中にエラーが発生しました:\n{str(e)}"))
    
    def export_all(self):
        """すべてのログ・情報を一括取得する（メインスレッド）"""
        # 別スレッドで実行して画面をブロックしないようにする
        thread = threading.Thread(target=self._export_all_thread)
        thread.daemon = True
        thread.start()
    
    def _export_all_thread(self):
        """すべてのログ・情報を一括取得する（バックグラウンド処理）"""
        try:
            # 処理中メッセージを表示
            self.root.after(0, lambda: self.show_progress_message("すべてのログ・情報を収集しています...\n少々お待ちください。"))
            
            # 現在の日時を取得してファイル名に使用
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
            
            # msinfo32コマンドを実行
            cmd = f'msinfo32 /report "{output_file}"'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            
            if result.returncode != 0:
                # エラーの場合は別の方法を試す
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
            
            # 処理中メッセージを閉じる
            self.root.after(0, self.hide_progress_message)
            
            # 完了メッセージを表示
            files_info = "出力ファイル:\n" + "\n".join([f"- {f}" for f in files_created]) + f"\n\n出力先:\n{self.output_folder}"
            self.root.after(0, lambda: self.show_completion_message(
                "完了",
                "すべてのログ・情報の出力が完了しました！",
                files_info
            ))
            
        except Exception as e:
            # 処理中メッセージを閉じる
            self.root.after(0, self.hide_progress_message)
            # エラーメッセージを表示
            self.root.after(0, lambda: messagebox.showerror("エラー", f"ログ・情報の出力中にエラーが発生しました:\n{str(e)}"))
    
    
    def export_msinfo(self):
        """msinfo32の情報をファイルに出力する（メインスレッド）"""
        # 別スレッドで実行して画面をブロックしないようにする
        thread = threading.Thread(target=self._export_msinfo_thread)
        thread.daemon = True
        thread.start()
    
    def _export_msinfo_thread(self):
        """msinfo32の情報をファイルに出力する（バックグラウンド処理）"""
        try:
            # 処理中メッセージを表示
            self.root.after(0, lambda: self.show_progress_message("システム情報を収集しています...\n少々お待ちください。"))
            
            # 現在の日時を取得してファイル名に使用
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(self.output_folder, f"SystemInfo_{timestamp}.txt")
            
            # msinfo32コマンドを実行
            cmd = f'msinfo32 /report "{output_file}"'
            
            # コマンドを実行
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            
            # 実行結果を確認
            if result.returncode == 0:
                # 処理中メッセージを閉じる
                self.root.after(0, self.hide_progress_message)
                # 完了メッセージを表示（改善版）
                files_info = f"出力ファイル:\n- {os.path.basename(output_file)}\n\n出力先:\n{self.output_folder}"
                self.root.after(0, lambda: self.show_completion_message(
                    "完了",
                    "システム情報の出力が完了しました！",
                    files_info
                ))
            else:
                # エラーの場合は別の方法を試す
                self.export_systeminfo_alternative(output_file)
                
        except Exception as e:
            # 処理中メッセージを閉じる
            self.root.after(0, self.hide_progress_message)
            # エラーメッセージを表示
            self.root.after(0, lambda: messagebox.showerror("エラー", f"システム情報の出力中にエラーが発生しました:\n{str(e)}"))
    
    def export_systeminfo_alternative(self, output_file):
        """代替方法でシステム情報を出力する"""
        try:
            # systeminfoコマンドを使用
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
                
                # 処理中メッセージを閉じる
                self.root.after(0, self.hide_progress_message)
                # 完了メッセージを表示（改善版）
                files_info = f"出力ファイル:\n- {os.path.basename(output_file)}\n\n出力先:\n{self.output_folder}"
                self.root.after(0, lambda: self.show_completion_message(
                    "完了",
                    "システム情報の出力が完了しました！",
                    files_info
                ))
            else:
                raise Exception("systeminfoコマンドも失敗しました")
                
        except Exception as e:
            # 処理中メッセージを閉じる
            self.root.after(0, self.hide_progress_message)
            # エラーメッセージを表示
            self.root.after(0, lambda: messagebox.showerror("エラー", f"システム情報の取得に失敗しました:\n{str(e)}"))
    
    def get_eventlog(self, log_name, output_file):
        """指定されたイベントログを取得してCSVに保存する"""
        # イベントログを開く
        hand = win32evtlog.OpenEventLog(None, log_name)
        
        # CSVファイルを作成
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            
            # ヘッダー行を書き込み
            writer.writerow([
                '日時', 'イベントID', 'レベル', 'ソース', 'メッセージ'
            ])
            
            # イベントログを読み取り（最新の1000件まで）
            events = win32evtlog.ReadEventLog(
                hand, 
                win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ,
                0
            )
            
            count = 0
            while events and count < 1000:  # 最大1000件まで
                for event in events:
                    # 日時の変換
                    time_generated = event.TimeGenerated.Format()
                    
                    # イベントレベルの判定
                    if event.EventType == win32con.EVENTLOG_ERROR_TYPE:
                        level = "エラー"
                    elif event.EventType == win32con.EVENTLOG_WARNING_TYPE:
                        level = "警告"
                    elif event.EventType == win32con.EVENTLOG_INFORMATION_TYPE:
                        level = "情報"
                    else:
                        level = "その他"
                    
                    # メッセージの取得
                    try:
                        message = win32evtlogutil.SafeFormatMessage(event, log_name)
                        if message is None:
                            message = "メッセージを取得できませんでした"
                    except:
                        message = "メッセージを取得できませんでした"
                    
                    # CSVに書き込み
                    writer.writerow([
                        time_generated,
                        event.EventID & 0xFFFF,  # イベントIDの下位16ビット
                        level,
                        event.SourceName,
                        message.replace('\n', ' ').replace('\r', '')  # 改行を除去
                    ])
                    
                    count += 1
                    if count >= 1000:
                        break
                
                # 次のイベントを読み取り
                if count < 1000:
                    events = win32evtlog.ReadEventLog(
                        hand, 
                        win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ,
                        0
                    )
                else:
                    break
        
        # イベントログを閉じる
        win32evtlog.CloseEventLog(hand)
    

    
    def run(self):
        """アプリケーションを実行する"""
        self.root.mainloop()

# メイン実行部分
if __name__ == "__main__":
    # 管理者権限の確認
    try:
        import ctypes
        if not ctypes.windll.shell32.IsUserAnAdmin():
            messagebox.showwarning(
                "権限不足", 
                "このプログラムは管理者権限で実行することを推奨します。\n" +
                "一部の機能が正常に動作しない可能性があります。"
            )
    except:
        pass
    
    # アプリケーションの起動
    app = ILCollector()
    app.run()