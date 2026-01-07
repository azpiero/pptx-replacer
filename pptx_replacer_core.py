#!/usr/bin/env python3
"""
PPTX画像置換 - コアエンジン

画像マッチング・置換のコアロジックを提供します。
GUIからもCLIからも利用可能です。
"""

import os
import shutil
import hashlib
import zipfile
import tempfile
from typing import Optional, List, Dict
from dataclasses import dataclass


# サポートする画像拡張子
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
    except Exception as e:
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
    
    # 置換用画像を読み込み
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


def batch_scan(
    folder_path: str,
    source_image_path: str,
    recursive: bool = True
) -> Dict[str, int]:
    """
    フォルダ内のPPTXファイルを一括スキャン
    
    Returns:
        Dict[pptx_path, match_count]
    """
    source_hash = calculate_file_hash(source_image_path)
    pptx_files = find_pptx_files(folder_path, recursive)
    
    results = {}
    for pptx_path in pptx_files:
        matches = scan_pptx_for_image(pptx_path, source_hash)
        match_count = sum(1 for m in matches if m.matched)
        results[pptx_path] = match_count
    
    return results


def batch_replace(
    folder_path: str,
    source_image_path: str,
    target_image_path: str,
    recursive: bool = True,
    output_folder: Optional[str] = None,
    backup: bool = True,
    progress_callback=None
) -> List[ReplaceResult]:
    """
    フォルダ内のPPTXファイルを一括置換
    
    Args:
        progress_callback: 進捗コールバック関数 (current, total, pptx_path)
    """
    source_hash = calculate_file_hash(source_image_path)
    
    # まずスキャン
    scan_results = batch_scan(folder_path, source_image_path, recursive)
    
    # マッチしたファイルのみ置換
    target_files = [(p, c) for p, c in scan_results.items() if c > 0]
    total = len(target_files)
    
    results = []
    for i, (pptx_path, _) in enumerate(target_files):
        # 出力パスの決定
        if output_folder:
            rel_path = os.path.relpath(pptx_path, folder_path)
            output_path = os.path.join(output_folder, rel_path)
        else:
            output_path = None
        
        result = replace_image_in_pptx(
            pptx_path, source_hash, target_image_path, output_path, backup
        )
        results.append(result)
        
        if progress_callback:
            progress_callback(i + 1, total, pptx_path)
    
    return results


# テスト用
if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="PPTX画像置換コアエンジン テスト")
    parser.add_argument("--scan", action="store_true", help="スキャンテスト")
    parser.add_argument("--folder", default=".", help="対象フォルダ")
    parser.add_argument("--source", help="入れ替え元画像")
    parser.add_argument("--target", help="入れ替え先画像")
    
    args = parser.parse_args()
    
    if args.scan and args.source:
        print(f"スキャン中: {args.folder}")
        results = batch_scan(args.folder, args.source)
        for path, count in results.items():
            status = "✓" if count > 0 else "−"
            print(f"  {status} {os.path.basename(path)}: {count}マッチ")
