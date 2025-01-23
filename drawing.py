import os
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
        "//ns:Default[@Extension='jpg']", namespaces=namespaces)

    # `PartName="/xl/drawings/drawing*.xml"` の `<Override>` 要素。
    drawing_overrides = root.xpath(
        "//ns:Override[starts-with(@PartName, '/xl/drawings/drawing')]",
        namespaces=namespaces)

    # print(jpg_defaults)
    # print(drawing_overrides)

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


def main():
    before_dir = Path('before')
    after_dir = Path('after')

    # フォルダをクリアする。
    if before_dir.exists():
        shutil.rmtree(before_dir)
    if after_dir.exists():
        shutil.rmtree(after_dir)

    # a.xlsx を before_dir に解凍する。
    with zipfile.ZipFile('a.xlsx', 'r') as zf:
        zf.extractall(str(before_dir))

    # a.xlsx を openpyxl で開いて a2.xlsx に保存する。
    wb = openpyxl.load_workbook('a.xlsx')
    wb.worksheets[0]['A1'].value = datetime.now()
    wb.save('a2.xlsx')
    wb.close()

    # a2.xlsx を after_dir に解凍する。
    with zipfile.ZipFile('a2.xlsx', 'r') as zf:
        zf.extractall(str(after_dir))

    # [Content_Types].xml 内の要素を復元 (before ⇒ after) する。
    restore_content_types(before_dir, after_dir)

    # xl/drawings/ フォルダを復元 (before ⇒ after) する。
    restore_drawings(before_dir, after_dir, 'xl/drawings')

    # xl/media/ フォルダを復元 (before ⇒ after) する。
    restore_drawings(before_dir, after_dir, 'xl/media/')

    # xl/worksheets/_rels/ フォルダを復元 (before ⇒ after) する。
    restore_drawings(before_dir, after_dir, 'xl/worksheets/_rels/')

    # TODO: xl/worksheets/sheet1.xml の内容を復元する。

    # TODO: xl/worksheets/sheet2.xml の内容を復元する。


if __name__ == '__main__':
    main()
