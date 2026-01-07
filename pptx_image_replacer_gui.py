#!/usr/bin/env python3
"""
PPTX画像一括置換ツール - GUI版

機能:
- 入れ替え元/入れ替え先画像の選択とプレビュー
- PPTXフォルダの探索（サブフォルダ対応）
- マッチング結果のプレビュー
- 一括置換実行（進捗バー付き）
- バックアップ機能
- 出力先フォルダ指定

必要ライブラリ:
- Pillow: pip install Pillow
"""

import os
import sys
import shutil
import hashlib
import zipfile
import tempfile
import threading
from typing import Optional, Dict, List
from dataclasses import dataclass

# =============================================================================
# Core Engine (画像マッチング・置換ロジック)
# =============================================================================

IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif',
                    '.wmf', '.emf', '.svg', '.wdp']


@dataclass
class MatchResult:
    """マッチング結果"""
    pptx_path: str
    internal_image_path: str
    internal_image_name: str
    matched: bool = False


@dataclass
class ReplaceResult:
    """置換結果"""
    pptx_path: str
    success: bool
    replaced_count: int
    message: str


def calculate_file_hash(filepath: str) -> str:
    """ファイルのMD5ハッシュを計算"""
    hash_md5 = hashlib.md5()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()


def calculate_bytes_hash(data: bytes) -> str:
    """バイトデータのMD5ハッシュを計算"""
    return hashlib.md5(data).hexdigest()


def find_pptx_files(directory: str, recursive: bool = True) -> List[str]:
    """指定ディレクトリ内のPPTXファイルを検索"""
    pptx_files = []
    if recursive:
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower().endswith('.pptx') and not file.startswith('~$'):
                    pptx_files.append(os.path.join(root, file))
    else:
        for file in os.listdir(directory):
            if file.lower().endswith('.pptx') and not file.startswith('~$'):
                pptx_files.append(os.path.join(directory, file))
    return sorted(pptx_files)


def scan_pptx_for_image(pptx_path: str, target_hash: str) -> List[MatchResult]:
    """PPTXファイル内で対象画像をスキャン"""
    results = []
    
    try:
        with zipfile.ZipFile(pptx_path, 'r') as zf:
            for name in zf.namelist():
                if name.startswith('ppt/media/'):
                    ext = os.path.splitext(name)[1].lower()
                    if ext in IMAGE_EXTENSIONS:
                        data = zf.read(name)
                        file_hash = calculate_bytes_hash(data)
                        matched = (file_hash == target_hash)
                        results.append(MatchResult(
                            pptx_path=pptx_path,
                            internal_image_path=name,
                            internal_image_name=os.path.basename(name),
                            matched=matched
                        ))
    except Exception:
        pass
    
    return results


def replace_image_in_pptx(
    pptx_path: str,
    target_hash: str,
    replacement_image_path: str,
    output_path: Optional[str] = None,
    backup: bool = True
) -> ReplaceResult:
    """PPTXファイル内の画像を置換"""
    
    if not os.path.exists(pptx_path):
        return ReplaceResult(pptx_path, False, 0, "ファイルが見つかりません")
    
    if not os.path.exists(replacement_image_path):
        return ReplaceResult(pptx_path, False, 0, "置換用画像が見つかりません")
    
    with open(replacement_image_path, 'rb') as f:
        replacement_data = f.read()
    
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_pptx = os.path.join(temp_dir, 'temp.pptx')
        replaced_count = 0
        
        try:
            with zipfile.ZipFile(pptx_path, 'r') as zf_in:
                with zipfile.ZipFile(temp_pptx, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                    for item in zf_in.namelist():
                        data = zf_in.read(item)
                        
                        if item.startswith('ppt/media/'):
                            ext = os.path.splitext(item)[1].lower()
                            if ext in IMAGE_EXTENSIONS:
                                file_hash = calculate_bytes_hash(data)
                                if file_hash == target_hash:
                                    data = replacement_data
                                    replaced_count += 1
                        
                        zf_out.writestr(item, data)
            
            if replaced_count > 0:
                if backup and output_path is None:
                    backup_path = pptx_path + '.backup'
                    if not os.path.exists(backup_path):
                        shutil.copy2(pptx_path, backup_path)
                
                final_path = output_path if output_path else pptx_path
                if output_path:
                    os.makedirs(os.path.dirname(output_path), exist_ok=True)
                shutil.copy2(temp_pptx, final_path)
                
                return ReplaceResult(pptx_path, True, replaced_count, f"{replaced_count}個の画像を置換")
            else:
                return ReplaceResult(pptx_path, True, 0, "マッチする画像なし")
                
        except Exception as e:
            return ReplaceResult(pptx_path, False, 0, f"エラー: {str(e)}")


# =============================================================================
# GUI Application
# =============================================================================

# Tkinterのインポート
try:
    import tkinter as tk
    from tkinter import ttk, filedialog
except ImportError:
    print("エラー: tkinterが見つかりません")
    print("Linuxの場合: sudo apt-get install python3-tk")
    print("Windowsの場合: Python再インストール時にtk/tclを含める")
    sys.exit(1)

# PILのインポート
try:
    from PIL import Image, ImageTk
except ImportError:
    print("エラー: Pillowライブラリが必要です")
    print("インストール: pip install Pillow")
    sys.exit(1)


class ImagePreviewFrame(ttk.LabelFrame):
    """画像プレビュー付きファイル選択フレーム"""
    
    def __init__(self, parent, title: str, **kwargs):
        super().__init__(parent, text=title, **kwargs)
        self.filepath = tk.StringVar()
        self.preview_size = (120, 90)
        self._create_widgets()
    
    def _create_widgets(self):
        path_frame = ttk.Frame(self)
        path_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.path_entry = ttk.Entry(path_frame, textvariable=self.filepath, width=40)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        browse_btn = ttk.Button(path_frame, text="参照...", command=self._browse_file)
        browse_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        preview_frame = ttk.Frame(self)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.preview_label = ttk.Label(preview_frame, text="画像を選択してください", 
                                        anchor=tk.CENTER, relief=tk.SUNKEN)
        self.preview_label.pack(fill=tk.BOTH, expand=True)
        
        self.info_label = ttk.Label(self, text="", foreground="gray")
        self.info_label.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        self.filepath.trace_add("write", self._on_path_change)
    
    def _browse_file(self):
        filetypes = [
            ("画像ファイル", "*.png *.jpg *.jpeg *.gif *.bmp"),
            ("PNG", "*.png"),
            ("JPEG", "*.jpg *.jpeg"),
            ("すべてのファイル", "*.*")
        ]
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            self.filepath.set(path)
    
    def _on_path_change(self, *args):
        self._update_preview()
    
    def _update_preview(self):
        path = self.filepath.get()
        
        if not path or not os.path.exists(path):
            self.preview_label.configure(image='', text="画像を選択してください")
            self.info_label.configure(text="")
            return
        
        try:
            img = Image.open(path)
            img.thumbnail(self.preview_size, Image.Resampling.LANCZOS)
            
            self.photo = ImageTk.PhotoImage(img)
            self.preview_label.configure(image=self.photo, text="")
            
            size = os.path.getsize(path)
            img_full = Image.open(path)
            info_text = f"{img_full.width}x{img_full.height} | {size:,} bytes | {os.path.basename(path)}"
            self.info_label.configure(text=info_text)
            
        except Exception as e:
            self.preview_label.configure(image='', text=f"プレビュー不可\n{str(e)}")
            self.info_label.configure(text="")
    
    def get_path(self) -> str:
        return self.filepath.get()
    
    def get_hash(self) -> Optional[str]:
        path = self.filepath.get()
        if path and os.path.exists(path):
            return calculate_file_hash(path)
        return None


class FolderSelectFrame(ttk.LabelFrame):
    """フォルダ選択フレーム"""
    
    def __init__(self, parent, title: str, **kwargs):
        super().__init__(parent, text=title, **kwargs)
        self.folderpath = tk.StringVar()
        self.recursive = tk.BooleanVar(value=True)
        self._create_widgets()
    
    def _create_widgets(self):
        path_frame = ttk.Frame(self)
        path_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.path_entry = ttk.Entry(path_frame, textvariable=self.folderpath, width=50)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        browse_btn = ttk.Button(path_frame, text="参照...", command=self._browse_folder)
        browse_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        option_frame = ttk.Frame(self)
        option_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        recursive_check = ttk.Checkbutton(option_frame, text="サブフォルダも検索", 
                                          variable=self.recursive)
        recursive_check.pack(side=tk.LEFT)
    
    def _browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folderpath.set(path)
    
    def get_path(self) -> str:
        return self.folderpath.get()
    
    def is_recursive(self) -> bool:
        return self.recursive.get()


class ResultTreeView(ttk.Frame):
    """結果表示ツリービュー"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self._create_widgets()
    
    def _create_widgets(self):
        columns = ("status", "file", "matches", "message")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", height=10)
        
        self.tree.heading("status", text="状態")
        self.tree.heading("file", text="ファイル")
        self.tree.heading("matches", text="マッチ数")
        self.tree.heading("message", text="メッセージ")
        
        self.tree.column("status", width=60, anchor=tk.CENTER)
        self.tree.column("file", width=300)
        self.tree.column("matches", width=80, anchor=tk.CENTER)
        self.tree.column("message", width=200)
        
        scrollbar_y = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
    
    def clear(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
    
    def add_scan_result(self, pptx_path: str, match_count: int):
        status = "✓" if match_count > 0 else "−"
        filename = os.path.basename(pptx_path)
        message = f"{match_count}個の画像がマッチ" if match_count > 0 else "マッチなし"
        self.tree.insert("", tk.END, values=(status, filename, match_count, message))
    
    def add_replace_result(self, result: ReplaceResult):
        if result.success and result.replaced_count > 0:
            status = "✓"
        elif result.success:
            status = "−"
        else:
            status = "✗"
        filename = os.path.basename(result.pptx_path)
        self.tree.insert("", tk.END, values=(status, filename, result.replaced_count, result.message))


class PPTXImageReplacerApp:
    """メインアプリケーション"""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PPTX画像一括置換ツール")
        self.root.geometry("750x700")
        self.root.minsize(650, 600)
        
        self.scan_results: Dict[str, int] = {}
        self.is_processing = False
        
        self._create_widgets()
        self._create_menu()
    
    def _create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ファイル", menu=file_menu)
        file_menu.add_command(label="終了", command=self.root.quit)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ヘルプ", menu=help_menu)
        help_menu.add_command(label="使い方", command=self._show_help)
        help_menu.add_command(label="バージョン情報", command=self._show_about)
    
    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 画像選択エリア
        image_frame = ttk.Frame(main_frame)
        image_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.source_image = ImagePreviewFrame(image_frame, "入れ替え元画像（検索対象）")
        self.source_image.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        arrow_label = ttk.Label(image_frame, text="→", font=("", 24))
        arrow_label.pack(side=tk.LEFT, padx=10)
        
        self.target_image = ImagePreviewFrame(image_frame, "入れ替え先画像（置換後）")
        self.target_image.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # フォルダ選択エリア
        self.folder_select = FolderSelectFrame(main_frame, "PPTX探索フォルダ")
        self.folder_select.pack(fill=tk.X, pady=(0, 10))
        
        # オプションエリア
        option_frame = ttk.LabelFrame(main_frame, text="オプション")
        option_frame.pack(fill=tk.X, pady=(0, 10))
        
        option_inner = ttk.Frame(option_frame)
        option_inner.pack(fill=tk.X, padx=5, pady=5)
        
        self.backup_var = tk.BooleanVar(value=True)
        backup_check = ttk.Checkbutton(option_inner, text="置換前にバックアップを作成", 
                                       variable=self.backup_var)
        backup_check.pack(side=tk.LEFT)
        
        self.separate_output_var = tk.BooleanVar(value=False)
        separate_check = ttk.Checkbutton(option_inner, text="別フォルダに出力:", 
                                         variable=self.separate_output_var,
                                         command=self._toggle_output_folder)
        separate_check.pack(side=tk.LEFT, padx=(20, 5))
        
        self.output_folder = tk.StringVar()
        self.output_entry = ttk.Entry(option_inner, textvariable=self.output_folder, 
                                      width=30, state=tk.DISABLED)
        self.output_entry.pack(side=tk.LEFT)
        
        self.output_browse_btn = ttk.Button(option_inner, text="参照...", 
                                            command=self._browse_output_folder,
                                            state=tk.DISABLED)
        self.output_browse_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # アクションボタン
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.scan_btn = ttk.Button(action_frame, text="スキャン実行", 
                                   command=self._start_scan, width=15)
        self.scan_btn.pack(side=tk.LEFT)
        
        self.replace_btn = ttk.Button(action_frame, text="置換実行", 
                                      command=self._start_replace, width=15,
                                      state=tk.DISABLED)
        self.replace_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(action_frame, variable=self.progress_var, 
                                            maximum=100, length=200)
        self.progress_bar.pack(side=tk.LEFT, padx=(20, 0), fill=tk.X, expand=True)
        
        self.progress_label = ttk.Label(action_frame, text="")
        self.progress_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # 結果表示エリア
        result_frame = ttk.LabelFrame(main_frame, text="結果")
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        self.result_tree = ResultTreeView(result_frame)
        self.result_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.summary_label = ttk.Label(result_frame, text="", font=("", 10))
        self.summary_label.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        # ステータスバー
        self.status_var = tk.StringVar(value="準備完了")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, 
                               relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
    
    def _toggle_output_folder(self):
        if self.separate_output_var.get():
            self.output_entry.configure(state=tk.NORMAL)
            self.output_browse_btn.configure(state=tk.NORMAL)
        else:
            self.output_entry.configure(state=tk.DISABLED)
            self.output_browse_btn.configure(state=tk.DISABLED)
    
    def _browse_output_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.output_folder.set(path)
    
    def _validate_inputs(self) -> bool:
        source_path = self.source_image.get_path()
        if not source_path or not os.path.exists(source_path):
            self.status_var.set("エラー: 入れ替え元画像を選択してください")
            print("エラー: 入れ替え元画像を選択してください")
            return False

        target_path = self.target_image.get_path()
        if not target_path or not os.path.exists(target_path):
            self.status_var.set("エラー: 入れ替え先画像を選択してください")
            print("エラー: 入れ替え先画像を選択してください")
            return False

        folder_path = self.folder_select.get_path()
        if not folder_path or not os.path.isdir(folder_path):
            self.status_var.set("エラー: PPTX探索フォルダを選択してください")
            print("エラー: PPTX探索フォルダを選択してください")
            return False

        if self.separate_output_var.get():
            output_path = self.output_folder.get()
            if not output_path:
                self.status_var.set("エラー: 出力先フォルダを指定してください")
                print("エラー: 出力先フォルダを指定してください")
                return False

        return True
    
    def _start_scan(self):
        if self.is_processing:
            return
        
        if not self._validate_inputs():
            return
        
        self.is_processing = True
        self.scan_btn.configure(state=tk.DISABLED)
        self.replace_btn.configure(state=tk.DISABLED)
        self.result_tree.clear()
        self.scan_results.clear()
        self.progress_var.set(0)
        self.status_var.set("スキャン中...")
        
        thread = threading.Thread(target=self._scan_worker)
        thread.daemon = True
        thread.start()
    
    def _scan_worker(self):
        try:
            source_hash = self.source_image.get_hash()
            folder_path = self.folder_select.get_path()
            recursive = self.folder_select.is_recursive()
            
            pptx_files = find_pptx_files(folder_path, recursive)
            total = len(pptx_files)
            
            if total == 0:
                self.root.after(0, lambda: self._scan_complete(0, 0))
                return
            
            matched_files = 0
            total_matches = 0
            
            for i, pptx_path in enumerate(pptx_files):
                results = scan_pptx_for_image(pptx_path, source_hash)
                match_count = sum(1 for r in results if r.matched)
                
                self.scan_results[pptx_path] = match_count
                
                if match_count > 0:
                    matched_files += 1
                    total_matches += match_count
                
                progress = (i + 1) / total * 100
                self.root.after(0, lambda p=progress, pp=pptx_path, mc=match_count: 
                               self._update_scan_progress(p, pp, mc))
            
            self.root.after(0, lambda: self._scan_complete(matched_files, total_matches))
            
        except Exception as e:
            self.root.after(0, lambda: self._scan_error(str(e)))
    
    def _update_scan_progress(self, progress: float, pptx_path: str, match_count: int):
        self.progress_var.set(progress)
        self.progress_label.configure(text=f"{int(progress)}%")
        self.result_tree.add_scan_result(pptx_path, match_count)
    
    def _scan_complete(self, matched_files: int, total_matches: int):
        self.is_processing = False
        self.scan_btn.configure(state=tk.NORMAL)
        
        total_files = len(self.scan_results)
        self.summary_label.configure(
            text=f"スキャン完了: {total_files}ファイル中 {matched_files}ファイルでマッチ（計{total_matches}画像）"
        )
        
        if matched_files > 0:
            self.replace_btn.configure(state=tk.NORMAL)
            self.status_var.set(f"スキャン完了 - {matched_files}ファイルで置換可能")
        else:
            self.status_var.set("スキャン完了 - マッチする画像なし")
    
    def _scan_error(self, error_msg: str):
        self.is_processing = False
        self.scan_btn.configure(state=tk.NORMAL)
        self.status_var.set(f"エラー発生: {error_msg}")
        print(f"スキャンエラー: {error_msg}")
    
    def _start_replace(self):
        if self.is_processing:
            return

        matched_files = sum(1 for c in self.scan_results.values() if c > 0)
        total_matches = sum(self.scan_results.values())

        msg = f"{matched_files}ファイル内の{total_matches}画像を置換します。"
        print(msg)
        self.status_var.set(msg)
        
        self.is_processing = True
        self.scan_btn.configure(state=tk.DISABLED)
        self.replace_btn.configure(state=tk.DISABLED)
        self.result_tree.clear()
        self.progress_var.set(0)
        self.status_var.set("置換中...")
        
        thread = threading.Thread(target=self._replace_worker)
        thread.daemon = True
        thread.start()
    
    def _replace_worker(self):
        try:
            source_hash = self.source_image.get_hash()
            target_path = self.target_image.get_path()
            backup = self.backup_var.get()
            
            use_separate_output = self.separate_output_var.get()
            output_base = self.output_folder.get() if use_separate_output else None
            source_base = self.folder_select.get_path()
            
            target_files = [(p, c) for p, c in self.scan_results.items() if c > 0]
            total = len(target_files)
            
            success_count = 0
            total_replaced = 0
            
            for i, (pptx_path, _) in enumerate(target_files):
                if output_base:
                    rel_path = os.path.relpath(pptx_path, source_base)
                    output_path = os.path.join(output_base, rel_path)
                else:
                    output_path = None
                
                result = replace_image_in_pptx(
                    pptx_path, source_hash, target_path, output_path, backup
                )
                
                if result.success and result.replaced_count > 0:
                    success_count += 1
                    total_replaced += result.replaced_count
                
                progress = (i + 1) / total * 100
                self.root.after(0, lambda p=progress, r=result: 
                               self._update_replace_progress(p, r))
            
            self.root.after(0, lambda: self._replace_complete(success_count, total_replaced))
            
        except Exception as e:
            self.root.after(0, lambda: self._replace_error(str(e)))
    
    def _update_replace_progress(self, progress: float, result: ReplaceResult):
        self.progress_var.set(progress)
        self.progress_label.configure(text=f"{int(progress)}%")
        self.result_tree.add_replace_result(result)
    
    def _replace_complete(self, success_count: int, total_replaced: int):
        self.is_processing = False
        self.scan_btn.configure(state=tk.NORMAL)
        self.scan_results.clear()

        self.summary_label.configure(
            text=f"置換完了: {success_count}ファイルで計{total_replaced}画像を置換しました"
        )
        self.status_var.set("置換完了")

        print(f"置換が完了しました。")
        print(f"成功: {success_count}ファイル")
        print(f"置換画像数: {total_replaced}")
    
    def _replace_error(self, error_msg: str):
        self.is_processing = False
        self.scan_btn.configure(state=tk.NORMAL)
        self.status_var.set(f"エラー発生: {error_msg}")
        print(f"置換エラー: {error_msg}")
    
    def _show_help(self):
        help_text = """【使い方】

1. 入れ替え元画像を選択
   PPTXファイル内で検索する画像を指定します。

2. 入れ替え先画像を選択
   置換後の新しい画像を指定します。

3. PPTX探索フォルダを選択
   PPTXファイルを検索するフォルダを指定します。

4. スキャン実行
   フォルダ内のPPTXファイルをスキャンし、
   マッチする画像を検索します。

5. 置換実行
   マッチした画像を一括で置換します。

【オプション】
・サブフォルダも検索: 下位フォルダも含めて検索
・バックアップ作成: 置換前に.backupファイルを作成
・別フォルダに出力: 元ファイルを変更せず別フォルダに出力
"""
        print(help_text)

    def _show_about(self):
        about_text = """PPTX画像一括置換ツール

Version 1.0.0

複数のPowerPointファイル内の特定画像を
一括で別の画像に置き換えるツールです。

(C) 2025"""
        print(about_text)


def main():
    root = tk.Tk()
    app = PPTXImageReplacerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
