#!/usr/bin/env python3
"""
PPTX内の画像ファイル一覧を取得するスクリプト（修正版）

- PPTX内部のファイル名（ppt/media/内）
- 元の画像ファイル名（XMLのname属性から取得可能な場合）
- 画像が使用されているスライド番号（presentation.xmlから正確に取得）
"""

import os
import sys
import zipfile
import hashlib
import re
from xml.etree import ElementTree as ET
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, field


# XML名前空間
NAMESPACES = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

# サポートする画像拡張子
IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif', 
                    '.wmf', '.emf', '.svg', '.wdp']


@dataclass
class ImageInfo:
    """画像情報を格納するデータクラス"""
    internal_path: str          # PPTX内部パス (ppt/media/image1.png)
    internal_name: str          # 内部ファイル名 (image1.png)
    original_name: Optional[str] = None  # 元のファイル名（取得できた場合）
    description: Optional[str] = None    # 説明文（alt text）
    size: int = 0               # ファイルサイズ
    md5_hash: str = ""          # MD5ハッシュ
    used_in_slides: List[int] = field(default_factory=list)  # 使用されているスライド番号
    shape_name: Optional[str] = None  # シェイプ名


def calculate_hash(data: bytes) -> str:
    """バイトデータのMD5ハッシュを計算"""
    return hashlib.md5(data).hexdigest()


def normalize_path(path: str) -> str:
    """パスを正規化（先頭のスラッシュを削除、バックスラッシュを変換）"""
    path = path.replace('\\', '/')
    if path.startswith('/'):
        path = path[1:]
    return path


def get_slide_order(zf: zipfile.ZipFile) -> List[Tuple[int, str]]:
    """
    presentation.xmlからスライドの順序を取得
    Returns: [(スライド番号, スライドファイルパス), ...]
    """
    slide_order = []
    
    try:
        # presentation.xml から rId の順序を取得
        pres_content = zf.read('ppt/presentation.xml').decode('utf-8')
        pres_root = ET.fromstring(pres_content)
        
        # sldIdLst から rId を順番に取得
        rid_order = []
        sld_id_lst = pres_root.find('.//p:sldIdLst', NAMESPACES)
        if sld_id_lst is not None:
            for sld_id in sld_id_lst.findall('p:sldId', NAMESPACES):
                rid = sld_id.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if rid:
                    rid_order.append(rid)
        
        # presentation.xml.rels から rId → ファイルパスの対応を取得
        rels_content = zf.read('ppt/_rels/presentation.xml.rels').decode('utf-8')
        rels_root = ET.fromstring(rels_content)
        
        rid_to_path = {}
        for rel in rels_root.findall('rel:Relationship', NAMESPACES):
            rid = rel.get('Id')
            target = rel.get('Target')
            rel_type = rel.get('Type', '')
            if 'slide' in rel_type.lower() and 'slideLayout' not in rel_type and 'slideMaster' not in rel_type:
                rid_to_path[rid] = normalize_path(target)
        
        # 順序通りにスライドをリストアップ
        for i, rid in enumerate(rid_order, 1):
            if rid in rid_to_path:
                slide_order.append((i, rid_to_path[rid]))
                
    except (KeyError, ET.ParseError) as e:
        print(f"Warning: スライド順序の取得に失敗: {e}", file=sys.stderr)
    
    return slide_order


def get_relationship_map(zf: zipfile.ZipFile, rels_path: str) -> Dict[str, str]:
    """リレーションシップファイルからrId→Targetのマッピングを取得"""
    rel_map = {}
    try:
        rels_content = zf.read(rels_path).decode('utf-8')
        # BOM除去
        if rels_content.startswith('\ufeff'):
            rels_content = rels_content[1:]
        root = ET.fromstring(rels_content)
        
        for rel in root.findall('rel:Relationship', NAMESPACES):
            rid = rel.get('Id')
            target = rel.get('Target')
            rel_type = rel.get('Type', '')
            # 画像関連のリレーションシップを取得
            if 'image' in rel_type.lower() or 'hdphoto' in rel_type.lower():
                rel_map[rid] = normalize_path(target)
    except (KeyError, ET.ParseError) as e:
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
                    target = rel_map[embed_id]
                    # 相対パスを絶対パスに変換
                    if target.startswith('..'):
                        internal_path = os.path.normpath(os.path.join(slide_dir, target))
                        internal_path = internal_path.replace('\\', '/')
                    elif target.startswith('ppt/'):
                        internal_path = target
                    else:
                        internal_path = f"ppt/{target}"
                    
                    internal_path = normalize_path(internal_path)
                    
                    # 画像情報を更新
                    if internal_path in image_info_map:
                        info = image_info_map[internal_path]
                        if slide_number not in info.used_in_slides:
                            info.used_in_slides.append(slide_number)
                        # 元のファイル名の候補を更新
                        if shape_name and not info.original_name:
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
    name_lower = name.lower()
    for ext in IMAGE_EXTENSIONS:
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
            normalized_name = normalize_path(name)
            if normalized_name.startswith('ppt/media/'):
                ext = os.path.splitext(name)[1].lower()
                if ext in IMAGE_EXTENSIONS:
                    data = zf.read(name)
                    image_info_map[normalized_name] = ImageInfo(
                        internal_path=normalized_name,
                        internal_name=os.path.basename(name),
                        size=len(data),
                        md5_hash=calculate_hash(data),
                    )
        
        # 2. presentation.xmlからスライドの順序を取得
        slide_order = get_slide_order(zf)
        
        # 3. 各スライドを解析
        for slide_number, slide_path in slide_order:
            try:
                extract_image_info_from_slide(zf, slide_path, slide_number, image_info_map)
            except Exception as e:
                print(f"Warning: スライド {slide_number} の解析に失敗: {e}", file=sys.stderr)
        
        # 4. スライドマスターやレイアウトからも情報を取得（スライド番号=0として扱う）
        for name in zf.namelist():
            normalized_name = normalize_path(name)
            if 'slideMaster' in name and name.endswith('.xml') and '_rels' not in name:
                extract_image_info_from_slide(zf, normalized_name, 0, image_info_map)
            elif 'slideLayout' in name and name.endswith('.xml') and '_rels' not in name:
                extract_image_info_from_slide(zf, normalized_name, 0, image_info_map)
    
    # リストに変換してソート
    images = list(image_info_map.values())
    images.sort(key=lambda x: x.internal_path)
    
    return images


def print_image_list(images: List[ImageInfo], pptx_path: str, verbose: bool = False) -> None:
    """画像情報を表示"""
    
    print(f"\n{'='*90}")
    print(f"PPTX: {pptx_path}")
    print(f"{'='*90}")
    print(f"画像数: {len(images)}")
    print(f"{'='*90}\n")
    
    if not images:
        print("画像が見つかりませんでした。")
        return
    
    # ヘッダー
    if verbose:
        print(f"{'No.':<4} {'内部ファイル名':<25} {'シェイプ名/説明':<30} {'サイズ':<10} {'スライド':<15} {'MD5ハッシュ'}")
        print("-" * 120)
    else:
        print(f"{'No.':<4} {'内部ファイル名':<25} {'シェイプ名/説明':<35} {'スライド'}")
        print("-" * 85)
    
    for i, img in enumerate(images, 1):
        # シェイプ名または説明を表示
        display_name = img.original_name or img.shape_name or img.description or "(なし)"
        # 長すぎる場合は切り詰め
        if len(display_name) > 30:
            display_name = display_name[:27] + "..."
        
        slides = ",".join(map(str, sorted(img.used_in_slides))) if img.used_in_slides else "(未使用)"
        
        if verbose:
            print(f"{i:<4} {img.internal_name:<25} {display_name:<30} {img.size:<10} {slides:<15} {img.md5_hash}")
        else:
            print(f"{i:<4} {img.internal_name:<25} {display_name:<35} {slides}")
    
    print()
    
    # サマリー
    with_original = [img for img in images if img.original_name]
    with_shape = [img for img in images if img.shape_name]
    with_desc = [img for img in images if img.description]
    unused = [img for img in images if not img.used_in_slides]
    
    print("--- サマリー ---")
    if with_original:
        print(f"元のファイル名が取得できた画像: {len(with_original)}/{len(images)}")
    print(f"シェイプ名あり: {len(with_shape)}/{len(images)}")
    print(f"説明(alt text)あり: {len(with_desc)}/{len(images)}")
    if unused:
        print(f"未使用画像（マスター/レイアウトのみ）: {len(unused)}")


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
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()