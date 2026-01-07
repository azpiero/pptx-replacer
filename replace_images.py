#!/usr/bin/env python3
"""
PowerPoint 画像一括置換スクリプト

複数のPPTXファイルに含まれる特定の画像を、別の画像に一括で置き換えます。
画像の識別は、ファイル名、ファイルハッシュ、またはファイルサイズで行えます。
"""

import os
import sys
import shutil
import hashlib
import argparse
import zipfile
import tempfile
from pathlib import Path
from typing import Optional, List, Dict, Tuple
import json


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


def get_image_info(filepath: str) -> Dict:
    """画像ファイルの情報を取得"""
    return {
        "filename": os.path.basename(filepath),
        "size": os.path.getsize(filepath),
        "hash": calculate_file_hash(filepath),
    }


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
    return pptx_files


def list_images_in_pptx(pptx_path: str) -> List[Dict]:
    """PPTXファイル内の画像一覧を取得"""
    images = []
    with zipfile.ZipFile(pptx_path, 'r') as zf:
        for name in zf.namelist():
            if name.startswith('ppt/media/'):
                ext = os.path.splitext(name)[1].lower()
                if ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.wmf', '.emf']:
                    data = zf.read(name)
                    images.append({
                        "path": name,
                        "filename": os.path.basename(name),
                        "size": len(data),
                        "hash": calculate_bytes_hash(data),
                    })
    return images


def replace_image_in_pptx(
    pptx_path: str,
    target_identifier: str,
    replacement_image_path: str,
    match_by: str = "hash",
    output_path: Optional[str] = None,
    backup: bool = True
) -> Tuple[bool, int, str]:
    """
    PPTXファイル内の画像を置換
    
    Args:
        pptx_path: 対象のPPTXファイルパス
        target_identifier: 置換対象の識別子（ハッシュ、ファイル名、またはサイズ）
        replacement_image_path: 置換用画像のパス
        match_by: マッチング方法 ("hash", "filename", "size")
        output_path: 出力先パス（Noneの場合は上書き）
        backup: バックアップを作成するか
    
    Returns:
        (成功フラグ, 置換数, メッセージ)
    """
    if not os.path.exists(pptx_path):
        return False, 0, f"ファイルが見つかりません: {pptx_path}"
    
    if not os.path.exists(replacement_image_path):
        return False, 0, f"置換用画像が見つかりません: {replacement_image_path}"
    
    # 置換用画像を読み込み
    with open(replacement_image_path, 'rb') as f:
        replacement_data = f.read()
    
    # 一時ディレクトリで作業
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_pptx = os.path.join(temp_dir, 'temp.pptx')
        replaced_count = 0
        
        with zipfile.ZipFile(pptx_path, 'r') as zf_in:
            with zipfile.ZipFile(temp_pptx, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf_in.namelist():
                    data = zf_in.read(item)
                    
                    # メディアファイルの場合、マッチングを確認
                    if item.startswith('ppt/media/'):
                        should_replace = False
                        
                        if match_by == "hash":
                            file_hash = calculate_bytes_hash(data)
                            should_replace = (file_hash == target_identifier)
                        elif match_by == "filename":
                            filename = os.path.basename(item)
                            should_replace = (filename == target_identifier)
                        elif match_by == "size":
                            should_replace = (len(data) == int(target_identifier))
                        
                        if should_replace:
                            # 画像を置換（拡張子を維持）
                            data = replacement_data
                            replaced_count += 1
                    
                    zf_out.writestr(item, data)
        
        if replaced_count > 0:
            # バックアップを作成
            if backup and output_path is None:
                backup_path = pptx_path + '.backup'
                if not os.path.exists(backup_path):
                    shutil.copy2(pptx_path, backup_path)
            
            # 出力
            final_path = output_path if output_path else pptx_path
            shutil.copy2(temp_pptx, final_path)
            return True, replaced_count, f"成功: {replaced_count}個の画像を置換しました"
        else:
            return True, 0, "マッチする画像が見つかりませんでした"


def batch_replace_images(
    directory: str,
    target_identifier: str,
    replacement_image_path: str,
    match_by: str = "hash",
    recursive: bool = True,
    output_dir: Optional[str] = None,
    backup: bool = True
) -> Dict[str, Tuple[bool, int, str]]:
    """
    複数のPPTXファイルで画像を一括置換
    
    Args:
        directory: PPTXファイルが含まれるディレクトリ
        target_identifier: 置換対象の識別子
        replacement_image_path: 置換用画像のパス
        match_by: マッチング方法
        recursive: サブディレクトリも検索するか
        output_dir: 出力先ディレクトリ（Noneの場合は上書き）
        backup: バックアップを作成するか
    
    Returns:
        ファイルパスと結果のマッピング
    """
    results = {}
    pptx_files = find_pptx_files(directory, recursive)
    
    print(f"\n見つかったPPTXファイル: {len(pptx_files)}件")
    print("-" * 50)
    
    for pptx_path in pptx_files:
        # 出力パスの決定
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            rel_path = os.path.relpath(pptx_path, directory)
            output_path = os.path.join(output_dir, rel_path)
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
        else:
            output_path = None
        
        success, count, message = replace_image_in_pptx(
            pptx_path,
            target_identifier,
            replacement_image_path,
            match_by,
            output_path,
            backup
        )
        results[pptx_path] = (success, count, message)
        
        status = "✓" if success and count > 0 else "○" if success else "✗"
        print(f"{status} {os.path.basename(pptx_path)}: {message}")
    
    return results


def analyze_target_image(image_path: str) -> None:
    """置換対象の画像を分析して識別子を表示"""
    if not os.path.exists(image_path):
        print(f"エラー: ファイルが見つかりません: {image_path}")
        return
    
    info = get_image_info(image_path)
    print("\n" + "=" * 50)
    print("画像分析結果")
    print("=" * 50)
    print(f"ファイル名: {info['filename']}")
    print(f"ファイルサイズ: {info['size']} bytes")
    print(f"MD5ハッシュ: {info['hash']}")
    print("\n使用例:")
    print(f"  --match-by hash --target {info['hash']}")
    print(f"  --match-by filename --target {info['filename']}")
    print(f"  --match-by size --target {info['size']}")
    print("=" * 50)


def scan_pptx_images(pptx_path: str) -> None:
    """PPTXファイル内の画像を一覧表示"""
    if not os.path.exists(pptx_path):
        print(f"エラー: ファイルが見つかりません: {pptx_path}")
        return
    
    images = list_images_in_pptx(pptx_path)
    print(f"\n{pptx_path} 内の画像一覧:")
    print("-" * 80)
    print(f"{'No.':<4} {'ファイル名':<25} {'サイズ':<12} {'MD5ハッシュ'}")
    print("-" * 80)
    
    for i, img in enumerate(images, 1):
        print(f"{i:<4} {img['filename']:<25} {img['size']:<12} {img['hash']}")
    
    print("-" * 80)
    print(f"合計: {len(images)}個の画像")


def scan_directory_images(directory: str, recursive: bool = True) -> None:
    """ディレクトリ内の全PPTXファイルの画像を分析"""
    pptx_files = find_pptx_files(directory, recursive)
    
    all_images = {}  # hash -> {info, files}
    
    for pptx_path in pptx_files:
        images = list_images_in_pptx(pptx_path)
        for img in images:
            h = img['hash']
            if h not in all_images:
                all_images[h] = {
                    "info": img,
                    "files": []
                }
            all_images[h]["files"].append(pptx_path)
    
    print(f"\n分析結果: {len(pptx_files)}個のPPTXファイル内のユニーク画像")
    print("=" * 100)
    print(f"{'ハッシュ':<34} {'ファイル名':<20} {'サイズ':<10} {'使用回数'}")
    print("=" * 100)
    
    for h, data in sorted(all_images.items(), key=lambda x: len(x[1]["files"]), reverse=True):
        info = data["info"]
        count = len(data["files"])
        print(f"{h:<34} {info['filename']:<20} {info['size']:<10} {count}ファイル")
    
    print("=" * 100)
    print(f"\nユニーク画像数: {len(all_images)}")
    
    # 複数ファイルで使用されている画像を詳細表示
    shared = {h: d for h, d in all_images.items() if len(d["files"]) > 1}
    if shared:
        print(f"\n複数ファイルで共通使用されている画像: {len(shared)}種類")
        for h, data in shared.items():
            print(f"\n  ハッシュ: {h}")
            print(f"  ファイル名: {data['info']['filename']}")
            print(f"  使用ファイル:")
            for f in data["files"]:
                print(f"    - {f}")


def main():
    parser = argparse.ArgumentParser(
        description="PowerPointファイル内の画像を一括置換するツール",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # 画像の分析（置換対象の識別子を取得）
  python replace_images.py --analyze /path/to/target_image.png
  
  # PPTXファイル内の画像一覧を表示
  python replace_images.py --scan /path/to/presentation.pptx
  
  # ディレクトリ内の全PPTXの画像を分析
  python replace_images.py --scan-dir /path/to/directory
  
  # ハッシュ値で画像を一括置換
  python replace_images.py --directory /path/to/pptx_folder \\
      --target abc123def456... --replacement /path/to/new_image.png
  
  # ファイル名で画像を置換（出力先を指定）
  python replace_images.py --directory /path/to/pptx_folder \\
      --match-by filename --target logo.png \\
      --replacement /path/to/new_logo.png --output-dir /path/to/output
        """
    )
    
    # モード選択
    mode_group = parser.add_mutually_exclusive_group(required=True)
    mode_group.add_argument("--analyze", metavar="IMAGE", help="画像ファイルを分析して識別子を表示")
    mode_group.add_argument("--scan", metavar="PPTX", help="PPTXファイル内の画像一覧を表示")
    mode_group.add_argument("--scan-dir", metavar="DIR", help="ディレクトリ内の全PPTXの画像を分析")
    mode_group.add_argument("--directory", "-d", metavar="DIR", help="置換対象のPPTXファイルがあるディレクトリ")
    
    # 置換オプション
    parser.add_argument("--target", "-t", help="置換対象の識別子（ハッシュ、ファイル名、またはサイズ）")
    parser.add_argument("--replacement", "-r", help="置換用画像のパス")
    parser.add_argument("--match-by", "-m", choices=["hash", "filename", "size"], default="hash",
                        help="マッチング方法（デフォルト: hash）")
    parser.add_argument("--output-dir", "-o", help="出力先ディレクトリ（指定しない場合は上書き）")
    parser.add_argument("--no-backup", action="store_true", help="バックアップを作成しない")
    parser.add_argument("--no-recursive", action="store_true", help="サブディレクトリを検索しない")
    
    args = parser.parse_args()
    
    # モードに応じた処理
    if args.analyze:
        analyze_target_image(args.analyze)
    elif args.scan:
        scan_pptx_images(args.scan)
    elif args.scan_dir:
        scan_directory_images(args.scan_dir, not args.no_recursive)
    elif args.directory:
        if not args.target or not args.replacement:
            parser.error("--directory モードでは --target と --replacement が必要です")
        
        results = batch_replace_images(
            args.directory,
            args.target,
            args.replacement,
            args.match_by,
            not args.no_recursive,
            args.output_dir,
            not args.no_backup
        )
        
        # 結果サマリー
        total = len(results)
        success = sum(1 for r in results.values() if r[0] and r[1] > 0)
        no_match = sum(1 for r in results.values() if r[0] and r[1] == 0)
        failed = sum(1 for r in results.values() if not r[0])
        total_replaced = sum(r[1] for r in results.values())
        
        print("\n" + "=" * 50)
        print("処理完了")
        print("=" * 50)
        print(f"処理ファイル数: {total}")
        print(f"置換成功: {success}ファイル（計{total_replaced}画像）")
        print(f"マッチなし: {no_match}ファイル")
        print(f"失敗: {failed}ファイル")


if __name__ == "__main__":
    main()
