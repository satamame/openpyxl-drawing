import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl.workbook.workbook import Workbook


def restore_content_types(before_dir: Path, after_dir: Path):
    '''[Content_Types].xml 内の要素を復元する。
    '''
    tree = etree.parse(before_dir / '[Content_Types].xml')
    root = tree.getroot()

    namespaces = {'ns': root.nsmap[None]}

    # `Content-Type` が "image/" で始まる `<Default>` 要素。
    image_defaults = root.xpath(
        "ns:Default[starts-with(@ContentType, 'image/')]",
        namespaces=namespaces)

    # `Extension="vml"` の `<Default>` 要素。
    vml_defaults = root.xpath(
        "ns:Default[@Extension='vml']", namespaces=namespaces)

    # `PartName="/xl/drawings/drawing*.xml"` の `<Override>` 要素。
    drawing_overrides = root.xpath(
        "ns:Override[starts-with(@PartName, '/xl/drawings/drawing')]",
        namespaces=namespaces)

    # 保存後の xml に追加する。
    tree = etree.parse(after_dir / '[Content_Types].xml')
    root = tree.getroot()

    for el in image_defaults:
        root.append(el)

    for el in vml_defaults:
        root.append(el)

    for el in drawing_overrides:
        root.append(el)

    # 保存する。
    tree = etree.ElementTree(root)
    tree.write(after_dir / '[Content_Types].xml', encoding='utf-8')


def restore_folder(before_dir: Path, after_dir: Path, folder2restore: str):
    '''folder2restore フォルダを復元する。
    '''
    src = before_dir / folder2restore
    dest = after_dir / folder2restore

    if not os.path.exists(dest):
        shutil.copytree(src, dest)


def restore_worksheets(before_dir: Path, after_dir: Path):
    src_dir = before_dir / 'xl/worksheets'
    dest_dir = after_dir / 'xl/worksheets'

    fname_ptn = re.compile(r'sheet[0-9]+\.xml')

    for f in src_dir.iterdir():
        if not (f.is_file() and fname_ptn.fullmatch(f.name)):
            continue

        f2 = dest_dir / f.name
        if not f2.exists():
            continue

        tree = etree.parse(f)
        root = tree.getroot()

        namespaces = {'ns': root.nsmap[None]}

        # `<drawing>` 要素。
        drawings = root.xpath("ns:drawing", namespaces=namespaces)

        # 保存後の xml に追加する。
        tree = etree.parse(f2)
        root = tree.getroot()

        for el in drawings:
            root.append(el)

        # 保存する。
        tree = etree.ElementTree(root)
        tree.write(f2, encoding='utf-8')


def save_with_drawings(
        wb: Workbook, src: Path, dest: Path, temp_dir_args={}):
    '''図形や画像を復元しつつ Workbook を保存する。

    Parameters
    ----------
    wb : Workbook
        保存する Workbook。
    src : Path
        復元する図形や画像の元となるブックファイルのパス。
    dest : Path
        Workbook の保存先となるブックファイルのパス。
    temp_dir_args : dict, default {}
        TemporaryDirectory を作る時のパラメータ。
    '''
    with tempfile.TemporaryDirectory(**temp_dir_args) as temp_dir:
        before_dir = Path(temp_dir) / 'before'
        after_dir = Path(temp_dir) / 'after'

        # src を before_dir に解凍する。
        with zipfile.ZipFile(src, 'r') as zf:
            zf.extractall(str(before_dir))

        # wb を dest に保存する。
        wb.save(dest)

        # dest を after_dir に解凍する。
        with zipfile.ZipFile(dest, 'r') as zf:
            zf.extractall(str(after_dir))

        # [Content_Types].xml 内の要素を復元 (before ⇒ after) する。
        # TODO: after_dir が存在するとコピーしない動作になっている。
        # TODO: 存在した場合でも特定のファイルやサブディレクトリをコピーするようにする。
        restore_content_types(before_dir, after_dir)

        # xl/drawings/ フォルダを復元 (before ⇒ after) する。
        restore_folder(before_dir, after_dir, 'xl/drawings/')

        # xl/media/ フォルダを復元 (before ⇒ after) する。
        restore_folder(before_dir, after_dir, 'xl/media/')

        # xl/worksheets/_rels/ フォルダを復元 (before ⇒ after) する。
        restore_folder(before_dir, after_dir, 'xl/worksheets/_rels/')

        # xl/worksheets/sheet*.xml の内容を復元 (before ⇒ after) する。
        restore_worksheets(before_dir, after_dir)

        # dest に圧縮しなおす。
        with zipfile.ZipFile(dest, 'w') as zf:
            for root, _, files in os.walk(after_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, after_dir)
                    zf.write(file_path, arcname)
