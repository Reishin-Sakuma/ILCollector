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
        self.root.geometry("400x300")
        self.root.resizable(False, False)
        
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
        
        # イベントログ出力ボタン
        eventlog_btn = tk.Button(
            self.root,
            text="イベントログをCSV出力",
            font=("Arial", 12),
            bg="lightblue",
            width=25,
            height=2,
            command=self.export_eventlogs
        )
        eventlog_btn.pack(pady=15)
        
        # msinfo32出力ボタン
        msinfo_btn = tk.Button(
            self.root,
            text="システム情報を出力",
            font=("Arial", 12),
            bg="lightgreen",
            width=25,
            height=2,
            command=self.export_msinfo
        )
        msinfo_btn.pack(pady=15)
        
        # 出力フォルダ表示
        folder_label = tk.Label(
            self.root,
            text=f"出力先: {self.output_folder}",
            font=("Arial", 9),
            fg="gray"
        )
        folder_label.pack(pady=10)
        
        # 終了ボタン
        exit_btn = tk.Button(
            self.root,
            text="終了",
            font=("Arial", 10),
            command=self.root.quit,
            bg="lightcoral"
        )
        exit_btn.pack(pady=10)
    
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
            
            # 完了メッセージを表示
            self.root.after(0, lambda: messagebox.showinfo(
                "完了", 
                f"イベントログの出力が完了しました！\n\n出力ファイル:\n- {os.path.basename(system_file)}\n- {os.path.basename(app_file)}\n\n出力先:\n{self.output_folder}"
            ))
            
        except Exception as e:
            # 処理中メッセージを閉じる
            self.root.after(0, self.hide_progress_message)
            # エラーメッセージを表示
            self.root.after(0, lambda: messagebox.showerror("エラー", f"イベントログの出力中にエラーが発生しました:\n{str(e)}"))
    
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
                # 完了メッセージを表示
                self.root.after(0, lambda: messagebox.showinfo(
                    "完了", 
                    f"システム情報の出力が完了しました！\n\n出力ファイル:\n{os.path.basename(output_file)}\n\n出力先:\n{self.output_folder}"
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
                # 完了メッセージを表示
                self.root.after(0, lambda: messagebox.showinfo(
                    "完了", 
                    f"システム情報の出力が完了しました！\n\n出力ファイル:\n{os.path.basename(output_file)}\n\n出力先:\n{self.output_folder}"
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