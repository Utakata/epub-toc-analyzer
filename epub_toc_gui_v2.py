#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
EPUB目次解析ツール GUI版 v2.0
バッチ処理、ドラッグ&ドロップ、プログレス表示対応
"""

import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import threading
import queue
from datetime import datetime

# ドラッグ&ドロップサポート
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_SUPPORT = True
except ImportError:
    DND_SUPPORT = False

# メインモジュールをインポート
try:
    from epubsplit_word_toc_v2 import SplitEpubWordTOC, BatchProcessor
    MODULE_AVAILABLE = True
except ImportError:
    MODULE_AVAILABLE = False

class EpubTocGUI:
    """EPUB目次解析ツール GUI版"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("📚 EPUB目次解析ツール v2.0 - Calibre互換版")
        self.root.geometry("800x700")
        self.root.configure(bg='#f0f0f0')
        
        # 変数
        self.processing = False
        self.queue = queue.Queue()
        self.batch_processor = None
        
        # スタイル設定
        self.setup_styles()
        
        # GUI作成
        self.create_widgets()
        
        # ドラッグ&ドロップ設定
        if DND_SUPPORT:
            self.setup_drag_drop()
        
        # キューチェック開始
        self.check_queue()
    
    def setup_styles(self):
        """スタイル設定"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # カスタムスタイル
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'), foreground='#2c5aa0')
        style.configure('Header.TLabel', font=('Arial', 12, 'bold'), foreground='#444444')
        style.configure('Status.TLabel', font=('Arial', 10), foreground='#666666')
        style.configure('Success.TLabel', font=('Arial', 10), foreground='#0a7029')
        style.configure('Error.TLabel', font=('Arial', 10), foreground='#d32f2f')
    
    def create_widgets(self):
        """ウィジェット作成"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # タイトル
        title_label = ttk.Label(main_frame, 
                               text="📚 EPUB目次解析ツール v2.0", 
                               style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # モード選択
        mode_frame = ttk.LabelFrame(main_frame, text="処理モード", padding="10")
        mode_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.mode_var = tk.StringVar(value="single")
        
        single_radio = ttk.Radiobutton(mode_frame, text="単一ファイル処理", 
                                      variable=self.mode_var, value="single",
                                      command=self.on_mode_change)
        single_radio.grid(row=0, column=0, padx=(0, 20))
        
        batch_radio = ttk.Radiobutton(mode_frame, text="バッチ処理（フォルダ内全EPUB）", 
                                     variable=self.mode_var, value="batch",
                                     command=self.on_mode_change)
        batch_radio.grid(row=0, column=1, padx=(0, 20))
        
        # ファイル/フォルダ選択
        input_frame = ttk.LabelFrame(main_frame, text="入力", padding="10")
        input_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.input_var = tk.StringVar()
        input_entry = ttk.Entry(input_frame, textvariable=self.input_var, width=60)
        input_entry.grid(row=0, column=0, padx=(0, 10))
        
        self.browse_button = ttk.Button(input_frame, text="参照", command=self.browse_input)
        self.browse_button.grid(row=0, column=1)
        
        # ドラッグ&ドロップエリア
        if DND_SUPPORT:
            self.drop_label = ttk.Label(input_frame, 
                                       text="📎 EPUBファイルまたはフォルダをドラッグ&ドロップ",
                                       style='Status.TLabel')
            self.drop_label.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        # 出力設定
        output_frame = ttk.LabelFrame(main_frame, text="出力設定", padding="10")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # 出力ディレクトリ
        ttk.Label(output_frame, text="出力ディレクトリ:").grid(row=0, column=0, sticky=tk.W)
        self.output_var = tk.StringVar(value=str(Path.cwd()))\n        output_entry = ttk.Entry(output_frame, textvariable=self.output_var, width=50)
        output_entry.grid(row=0, column=1, padx=(10, 10))
        
        output_browse = ttk.Button(output_frame, text="参照", command=self.browse_output)
        output_browse.grid(row=0, column=2)
        
        # 出力形式
        ttk.Label(output_frame, text="出力形式:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.format_var = tk.StringVar(value="both")
        format_combo = ttk.Combobox(output_frame, textvariable=self.format_var, 
                                   values=["both", "text", "word"], state="readonly", width=15)
        format_combo.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=(10, 0))
        
        # バッチ処理設定
        self.batch_frame = ttk.LabelFrame(main_frame, text="バッチ処理設定", padding="10")
        self.batch_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        ttk.Label(self.batch_frame, text="並列実行数:").grid(row=0, column=0, sticky=tk.W)
        self.workers_var = tk.StringVar(value="4")
        workers_spin = ttk.Spinbox(self.batch_frame, from_=1, to=8, 
                                  textvariable=self.workers_var, width=10)
        workers_spin.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        
        recursive_var = tk.BooleanVar(value=True)
        recursive_check = ttk.Checkbutton(self.batch_frame, text="サブフォルダも検索", 
                                         variable=recursive_var)
        recursive_check.grid(row=0, column=2, padx=(20, 0))
        
        # 実行ボタン
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=(0, 15))
        
        self.process_button = ttk.Button(button_frame, text="🚀 解析開始", 
                                        command=self.start_processing, 
                                        style='Accent.TButton')
        self.process_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_button = ttk.Button(button_frame, text="⏹ 停止", 
                                     command=self.stop_processing, 
                                     state='disabled')
        self.stop_button.pack(side=tk.LEFT)
        
        # プログレスバー
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # ステータス
        self.status_var = tk.StringVar(value="Ready")\n        status_label = ttk.Label(main_frame, textvariable=self.status_var, style='Status.TLabel')
        status_label.grid(row=7, column=0, columnspan=3)
        
        # ログエリア
        log_frame = ttk.LabelFrame(main_frame, text="処理ログ", padding="10")
        log_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(15, 0))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, width=80, height=15, 
                                                 font=('Consolas', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # グリッドの重みを設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(8, weight=1)
        
        # 初期状態設定
        self.on_mode_change()
    
    def setup_drag_drop(self):
        """ドラッグ&ドロップの設定"""
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)
    
    def on_drop(self, event):
        """ドラッグ&ドロップ処理"""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = Path(files[0])
            if file_path.is_file() and file_path.suffix.lower() == '.epub':
                self.input_var.set(str(file_path))
                self.mode_var.set("single")
                self.on_mode_change()
            elif file_path.is_dir():
                self.input_var.set(str(file_path))
                self.mode_var.set("batch")
                self.on_mode_change()
            else:
                messagebox.showwarning("警告", "EPUBファイルまたはフォルダを選択してください")
    
    def on_mode_change(self):
        """モード変更時の処理"""
        if self.mode_var.get() == "single":
            self.batch_frame.grid_remove()
            self.browse_button.configure(text="ファイル選択")
        else:
            self.batch_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
            self.browse_button.configure(text="フォルダ選択")
    
    def browse_input(self):
        """入力ファイル/フォルダの選択"""
        if self.mode_var.get() == "single":
            file_path = filedialog.askopenfilename(
                title="EPUBファイルを選択",
                filetypes=[("EPUB files", "*.epub"), ("All files", "*.*")]
            )
            if file_path:
                self.input_var.set(file_path)
        else:
            folder_path = filedialog.askdirectory(title="EPUBファイルが含まれるフォルダを選択")
            if folder_path:
                self.input_var.set(folder_path)
    
    def browse_output(self):
        """出力フォルダの選択"""
        folder_path = filedialog.askdirectory(title="出力フォルダを選択")
        if folder_path:
            self.output_var.set(folder_path)
    
    def start_processing(self):
        """処理開始"""
        if not MODULE_AVAILABLE:
            messagebox.showerror("エラー", "必要なモジュールがインストールされていません")
            return
        
        input_path = self.input_var.get().strip()
        if not input_path:
            messagebox.showwarning("警告", "入力ファイル/フォルダを選択してください")
            return
        
        if not Path(input_path).exists():
            messagebox.showerror("エラー", "指定されたパスが存在しません")
            return
        
        # UI状態更新
        self.processing = True
        self.process_button.configure(state='disabled')
        self.stop_button.configure(state='normal')
        self.progress.configure(mode='indeterminate')
        self.progress.start()
        self.status_var.set("処理中...")
        
        # ログクリア
        self.log_text.delete(1.0, tk.END)
        self.log(f"🚀 処理開始: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 別スレッドで処理実行
        thread = threading.Thread(target=self.process_files, daemon=True)
        thread.start()
    
    def process_files(self):
        """ファイル処理（別スレッド）"""
        try:
            input_path = Path(self.input_var.get())
            output_dir = self.output_var.get()
            format_type = self.format_var.get()
            
            if self.mode_var.get() == "single":
                # 単一ファイル処理
                self.queue.put(("log", f"📚 ファイル処理: {input_path.name}"))
                
                with open(input_path, 'rb') as f:
                    epub_splitter = SplitEpubWordTOC(f)
                    results = epub_splitter.generate_word_toc_output(
                        output_dir=output_dir,
                        format_type=format_type
                    )
                
                self.queue.put(("log", "✅ 処理完了!"))
                for format_name, file_path in results.items():
                    self.queue.put(("log", f"   📄 {format_name}: {Path(file_path).name}"))
                
                self.queue.put(("complete", "単一ファイルの処理が完了しました"))
                
            else:
                # バッチ処理
                workers = int(self.workers_var.get())
                self.queue.put(("log", f"🔄 バッチ処理開始: {input_path}"))
                self.queue.put(("log", f"⚙️ 並列実行数: {workers}"))
                
                self.batch_processor = BatchProcessor(max_workers=workers)
                results = self.batch_processor.process_directory(
                    str(input_path),
                    output_dir=output_dir,
                    format_type=format_type
                )
                
                success_count = len([r for r in results if r['status'] == 'success'])
                error_count = len([r for r in results if r['status'] == 'error'])
                
                self.queue.put(("log", f"📊 バッチ処理完了:"))
                self.queue.put(("log", f"   ✅ 成功: {success_count}ファイル"))
                self.queue.put(("log", f"   ❌ エラー: {error_count}ファイル"))
                
                if error_count > 0:
                    self.queue.put(("log", f"❌ エラー詳細:"))
                    for error in self.batch_processor.errors[:3]:
                        self.queue.put(("log", f"   - {error}"))
                
                self.queue.put(("complete", f"バッチ処理が完了しました（{success_count}件成功）"))
                
        except Exception as e:
            self.queue.put(("error", f"処理エラー: {str(e)}"))
        finally:
            self.queue.put(("finished", None))
    
    def stop_processing(self):
        """処理停止"""
        self.processing = False
        self.queue.put(("log", "⏹ 処理を停止しました"))
        self.queue.put(("finished", None))
    
    def check_queue(self):
        """キューチェック"""
        try:
            while True:
                msg_type, msg_data = self.queue.get_nowait()
                
                if msg_type == "log":
                    self.log(msg_data)
                elif msg_type == "complete":
                    self.status_var.set(msg_data)
                    messagebox.showinfo("完了", msg_data)
                elif msg_type == "error":
                    self.status_var.set(msg_data)
                    messagebox.showerror("エラー", msg_data)
                elif msg_type == "finished":
                    self.finish_processing()
                    break
                    
        except queue.Empty:
            pass
        
        # 100ms後に再チェック
        self.root.after(100, self.check_queue)
    
    def log(self, message):
        """ログ出力"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\\n"
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def finish_processing(self):
        """処理終了時の後処理"""
        self.processing = False
        self.process_button.configure(state='normal')
        self.stop_button.configure(state='disabled')
        self.progress.stop()
        self.progress.configure(mode='determinate', value=0)
        
        if not self.status_var.get().startswith("✅") and not self.status_var.get().startswith("❌"):
            self.status_var.set("Ready")

def main():
    """メイン関数"""
    # Tkinterルートウィンドウを作成
    if DND_SUPPORT:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        print("⚠️ tkinterdnd2がインストールされていません。ドラッグ&ドロップは無効です。")
    
    try:
        # GUI作成
        app = EpubTocGUI(root)
        
        # ウィンドウを中央に配置
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
        y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
        root.geometry(f"+{x}+{y}")
        
        # メインループ開始
        root.mainloop()
        
    except Exception as e:
        messagebox.showerror("起動エラー", f"アプリケーションの起動に失敗しました:\\n{str(e)}")

if __name__ == "__main__":
    main()
