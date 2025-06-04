# ILCollector

## 概要

**ILCollector** は、Windows Server の「イベントログ」と「システム情報」を簡単に収集・出力できるGUIツールです。  
System/ApplicationイベントログをCSV形式で、システム情報をテキスト形式で取得できます。  
管理者権限での実行を推奨しますが、権限がない場合も一部機能は利用可能です。

---

## 主な機能

- **イベントログ出力**  
  System・ApplicationログをCSVファイルとして出力

- **システム情報出力**  
  CPU・メモリ・OS情報などの詳細情報をテキストで出力

- **すべて一括取得**  
  上記の処理をまとめて一括実行

- **出力フォルダをエクスプローラーで開く**  
  ワンクリックで出力先フォルダを開けます

---

## 使い方

1. **Python環境の準備**  
   - Python 3.8 以上を推奨
   - 必要なパッケージをインストール  
     ```
     pip install pywin32
     ```

2. **起動方法**  
   コマンドプロンプトで本プログラムのあるフォルダに移動し、以下を実行  
   ```
   python ILCollector.py
   ```

3. **管理者権限について**  
   - 管理者権限で実行していない場合、起動時に警告ダイアログが表示されます。
   - 「続行」→ 権限がないまま起動  
   - 「終了」→ プログラムを終了

4. **操作方法**  
   - 起動後、GUI画面から各ボタンをクリックして機能を実行してください。
   - 出力ファイルは自動でタイムスタンプ付きのフォルダに保存されます。

---

## 出力ファイル

- `System_EventLog_YYYYMMDD_HHMMSS.csv`  
- `Application_EventLog_YYYYMMDD_HHMMSS.csv`  
- `SystemInfo_YYYYMMDD_HHMMSS.txt`  

出力先は画面下部に表示されます。

---

## 必要なモジュール

- `tkinter`
- `pywin32`（`win32evtlog`, `win32evtlogutil`, `win32con` など）
- `csv`
- `os`
- `subprocess`
- `datetime`
- `threading`
- `sys`

---

## 注意事項

- 一部のイベントログやシステム情報は管理者権限がないと取得できません。
- Windows専用ツールです。
- 出力ファイルはUTF-8で保存されます。

---

## ライセンス

本ツールは社内利用・検証目的で作成されています。