import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl
from lxml import etree


def restore_content_types(before_dir: Path, after_dir: Path):
    '''[Content_Types].xml 内の要素を復元する。
    '''
    tree = etree.parse(before_dir / '[Content_Types].xml')
    root = tree.getroot()

    namespaces = {'ns': root.nsmap[None]}

    # `Extension="jpg"` の `<Default>` 要素。
    jpg_defaults = root.xpath(
        "ns:Default[@Extension='jpg']", namespaces=namespaces)

    # `PartName="/xl/drawings/drawing*.xml"` の `<Override>` 要素。
    drawing_overrides = root.xpath(
        "ns:Override[starts-with(@PartName, '/xl/drawings/drawing')]",
        namespaces=namespaces)

    # 保存後の xml に追加する。
    tree = etree.parse(after_dir / '[Content_Types].xml')
    root = tree.getroot()

    for el in jpg_defaults:
        root.append(el)

    for el in drawing_overrides:
        root.append(el)

    # 保存する。
    tree = etree.ElementTree(root)
    tree.write(after_dir / '[Content_Types].xml', encoding='utf-8')


def restore_drawings(before_dir: Path, after_dir: Path, folder2restore: str):
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


def main():
    before_dir = Path('temp/before')
    after_dir = Path('temp/after')
    open_file = Path('a.xlsx')
    save_file = Path('temp/a2.xlsx')

    # フォルダをクリアする。
    if before_dir.exists():
        shutil.rmtree(before_dir)
    if after_dir.exists():
        shutil.rmtree(after_dir)

    # a.xlsx を before_dir に解凍する。
    with zipfile.ZipFile(open_file, 'r') as zf:
        zf.extractall(str(before_dir))

    # a.xlsx を openpyxl で開いて a2.xlsx に保存する。
    wb = openpyxl.load_workbook(open_file)
    wb.worksheets[0]['A1'].value = datetime.now()
    wb.save(save_file)
    wb.close()

    # a2.xlsx を after_dir に解凍する。
    with zipfile.ZipFile(save_file, 'r') as zf:
        zf.extractall(str(after_dir))

    # [Content_Types].xml 内の要素を復元 (before ⇒ after) する。
    restore_content_types(before_dir, after_dir)

    # xl/drawings/ フォルダを復元 (before ⇒ after) する。
    restore_drawings(before_dir, after_dir, 'xl/drawings/')

    # xl/media/ フォルダを復元 (before ⇒ after) する。
    restore_drawings(before_dir, after_dir, 'xl/media/')

    # xl/worksheets/_rels/ フォルダを復元 (before ⇒ after) する。
    restore_drawings(before_dir, after_dir, 'xl/worksheets/_rels/')

    # xl/worksheets/sheet*.xml の内容を復元 (before ⇒ after) する。
    restore_worksheets(before_dir, after_dir)

    # save_file に圧縮しなおす。
    with zipfile.ZipFile(save_file, 'w') as zf:
        for root, _, files in os.walk(after_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, after_dir)
                zf.write(file_path, arcname)


if __name__ == '__main__':
    main()
