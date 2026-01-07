#!/usr/bin/env python3
"""
PPTX内の画像ファイル一覧を取得するスクリプト

- PPTX内部のファイル名（ppt/media/内）
- 元の画像ファイル名（XMLのname属性から取得可能な場合）
- 画像が使用されているスライド番号
"""

import os
import sys
import zipfile
import hashlib
import re
from xml.etree import ElementTree as ET
from typing import Dict, List, Optional
from dataclasses import dataclass


# XML名前空間
NAMESPACES = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}


@dataclass
class ImageInfo:
    """画像情報を格納するデータクラス"""
    internal_path: str          # PPTX内部パス (ppt/media/image1.png)
    internal_name: str          # 内部ファイル名 (image1.png)
    original_name: Optional[str] = None  # 元のファイル名（取得できた場合）
    description: Optional[str] = None    # 説明文（alt text）
    size: int = 0               # ファイルサイズ
    md5_hash: str = ""          # MD5ハッシュ
    used_in_slides: List[int] = None  # 使用されているスライド番号
    shape_name: Optional[str] = None  # シェイプ名

    def __post_init__(self):
        if self.used_in_slides is None:
            self.used_in_slides = []


def calculate_hash(data: bytes) -> str:
    """バイトデータのMD5ハッシュを計算"""
    return hashlib.md5(data).hexdigest()


def get_relationship_map(zf: zipfile.ZipFile, rels_path: str) -> Dict[str, str]:
    """リレーションシップファイルからrId→Targetのマッピングを取得"""
    rel_map = {}
    try:
        rels_content = zf.read(rels_path).decode('utf-8')
        root = ET.fromstring(rels_content)
        for rel in root.findall('rel:Relationship', NAMESPACES):
            rid = rel.get('Id')
            target = rel.get('Target')
            rel_type = rel.get('Type', '')
            if 'image' in rel_type.lower():
                rel_map[rid] = target
    except (KeyError, ET.ParseError):
        pass
    return rel_map


def extract_image_info_from_slide(
    zf: zipfile.ZipFile, 
    slide_path: str, 
    slide_number: int,
    image_info_map: Dict[str, ImageInfo]
) -> None:
    """スライドXMLから画像情報を抽出"""
    
    # リレーションシップファイルのパス
    slide_dir = os.path.dirname(slide_path)
    slide_name = os.path.basename(slide_path)
    rels_path = f"{slide_dir}/_rels/{slide_name}.rels"
    
    # rId→画像パスのマッピングを取得
    rel_map = get_relationship_map(zf, rels_path)
    
    try:
        slide_content = zf.read(slide_path).decode('utf-8')
        root = ET.fromstring(slide_content)
        
        # p:pic要素（画像）を検索
        for pic in root.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}pic'):
            # nvPicPr > cNvPr から名前と説明を取得
            cNvPr = pic.find('.//p:nvPicPr/p:cNvPr', NAMESPACES)
            shape_name = None
            description = None
            
            if cNvPr is not None:
                shape_name = cNvPr.get('name')
                description = cNvPr.get('descr')
            
            # blipFill > blip から画像参照(rId)を取得
            blip = pic.find('.//p:blipFill/a:blip', NAMESPACES)
            if blip is not None:
                embed_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                
                if embed_id and embed_id in rel_map:
                    # 相対パスを絶対パスに変換
                    target = rel_map[embed_id]
                    if target.startswith('..'):
                        # ../media/image1.png → ppt/media/image1.png
                        internal_path = os.path.normpath(os.path.join(slide_dir, target))
                        internal_path = internal_path.replace('\\', '/')
                    else:
                        internal_path = f"ppt/{target}"
                    
                    # 画像情報を更新または追加
                    if internal_path in image_info_map:
                        info = image_info_map[internal_path]
                        if slide_number not in info.used_in_slides:
                            info.used_in_slides.append(slide_number)
                        # 元のファイル名の候補を更新（より良い情報があれば）
                        if shape_name and not info.original_name:
                            # shape_nameが画像ファイル名っぽければ採用
                            if is_likely_filename(shape_name):
                                info.original_name = shape_name
                        if not info.shape_name:
                            info.shape_name = shape_name
                        if not info.description:
                            info.description = description
                            
    except (KeyError, ET.ParseError) as e:
        pass


def is_likely_filename(name: str) -> bool:
    """文字列がファイル名っぽいかどうかを判定"""
    if not name:
        return False
    # 画像拡張子を含むか
    image_extensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.wmf', '.emf', '.svg']
    name_lower = name.lower()
    for ext in image_extensions:
        if name_lower.endswith(ext):
            return True
    return False


def list_images_in_pptx(pptx_path: str) -> List[ImageInfo]:
    """PPTXファイル内の全画像情報を取得"""
    
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(f"ファイルが見つかりません: {pptx_path}")
    
    image_info_map: Dict[str, ImageInfo] = {}
    
    with zipfile.ZipFile(pptx_path, 'r') as zf:
        # 1. まずppt/media/内の全画像を取得
        for name in zf.namelist():
            if name.startswith('ppt/media/'):
                ext = os.path.splitext(name)[1].lower()
                if ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.wmf', '.emf', '.svg']:
                    data = zf.read(name)
                    image_info_map[name] = ImageInfo(
                        internal_path=name,
                        internal_name=os.path.basename(name),
                        size=len(data),
                        md5_hash=calculate_hash(data),
                    )
        
        # 2. 各スライドを解析して画像の使用状況と元ファイル名を取得
        slide_pattern = re.compile(r'ppt/slides/slide(\d+)\.xml$')
        slide_files = []
        
        for name in zf.namelist():
            match = slide_pattern.match(name)
            if match:
                slide_number = int(match.group(1))
                slide_files.append((slide_number, name))
        
        # スライド番号順にソート
        slide_files.sort(key=lambda x: x[0])
        
        for slide_number, slide_path in slide_files:
            extract_image_info_from_slide(zf, slide_path, slide_number, image_info_map)
        
        # 3. スライドマスターやレイアウトからも情報を取得
        for name in zf.namelist():
            if 'slideMaster' in name and name.endswith('.xml') and '_rels' not in name:
                extract_image_info_from_slide(zf, name, 0, image_info_map)
            elif 'slideLayout' in name and name.endswith('.xml') and '_rels' not in name:
                extract_image_info_from_slide(zf, name, 0, image_info_map)
    
    # リストに変換してソート
    images = list(image_info_map.values())
    images.sort(key=lambda x: x.internal_path)
    
    return images


def print_image_list(images: List[ImageInfo], pptx_path: str, verbose: bool = False) -> None:
    """画像情報を表示"""
    
    print(f"\n{'='*80}")
    print(f"PPTX: {pptx_path}")
    print(f"{'='*80}")
    print(f"画像数: {len(images)}")
    print(f"{'='*80}\n")
    
    if not images:
        print("画像が見つかりませんでした。")
        return
    
    # ヘッダー
    if verbose:
        print(f"{'No.':<4} {'内部ファイル名':<20} {'元のファイル名':<25} {'サイズ':<10} {'スライド':<12} {'MD5ハッシュ'}")
        print("-" * 110)
    else:
        print(f"{'No.':<4} {'内部ファイル名':<20} {'元のファイル名':<30} {'スライド'}")
        print("-" * 75)
    
    for i, img in enumerate(images, 1):
        original = img.original_name or img.shape_name or "(取得不可)"
        slides = ",".join(map(str, sorted(img.used_in_slides))) if img.used_in_slides else "(未使用)"
        
        if verbose:
            print(f"{i:<4} {img.internal_name:<20} {original:<25} {img.size:<10} {slides:<12} {img.md5_hash}")
        else:
            print(f"{i:<4} {img.internal_name:<20} {original:<30} {slides}")
    
    print()
    
    # 元ファイル名が取得できた画像のサマリー
    with_original = [img for img in images if img.original_name]
    if with_original:
        print(f"元のファイル名が取得できた画像: {len(with_original)}/{len(images)}")
    else:
        print("※ 元のファイル名は取得できませんでした（PowerPointが保存時に破棄した可能性）")


def export_to_json(images: List[ImageInfo], output_path: str) -> None:
    """画像情報をJSONファイルに出力"""
    import json
    
    data = []
    for img in images:
        data.append({
            "internal_path": img.internal_path,
            "internal_name": img.internal_name,
            "original_name": img.original_name,
            "shape_name": img.shape_name,
            "description": img.description,
            "size": img.size,
            "md5_hash": img.md5_hash,
            "used_in_slides": img.used_in_slides,
        })
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"JSONファイルに出力しました: {output_path}")


def main():
    import argparse
    
    parser = argparse.ArgumentParser(
        description="PPTXファイル内の画像ファイル一覧を取得",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python list_pptx_images.py presentation.pptx
  python list_pptx_images.py presentation.pptx -v
  python list_pptx_images.py presentation.pptx --json output.json
        """
    )
    
    parser.add_argument("pptx_file", help="対象のPPTXファイル")
    parser.add_argument("-v", "--verbose", action="store_true", 
                        help="詳細情報（サイズ、ハッシュ）も表示")
    parser.add_argument("--json", metavar="FILE", 
                        help="結果をJSONファイルに出力")
    
    args = parser.parse_args()
    
    try:
        images = list_images_in_pptx(args.pptx_file)
        print_image_list(images, args.pptx_file, args.verbose)
        
        if args.json:
            export_to_json(images, args.json)
            
    except FileNotFoundError as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)
    except zipfile.BadZipFile:
        print(f"エラー: 有効なPPTXファイルではありません: {args.pptx_file}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()