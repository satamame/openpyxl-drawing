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
    before_tree = etree.parse(before_dir / '[Content_Types].xml')
    before_root = before_tree.getroot()
    after_tree = etree.parse(after_dir / '[Content_Types].xml')
    after_root = after_tree.getroot()

    # openpyxl によって保存された Override 要素
    overrides = after_root.findall(".//{*}Override")
    after_overrides = {
        child.get("PartName") for child in overrides if child.get("PartName")
    }

    # openpyxl によって保存された Default 要素
    defaults = after_root.findall(".//{*}Default")
    after_defaults = {
        child.get("Extension") for child in defaults if child.get("Extension")
    }

    # Add missing <Override> tags from original to modified
    cmt_ptn = re.compile(r'/xl/comments\d+\.xml')
    for child in before_root.findall(".//{*}Override"):
        part_name = child.get("PartName")
        if not part_name:
            continue
        # xl/calcChain.xml は Excel が復元するので、復元しない。
        if part_name == '/xl/calcChain.xml':
            continue
        # xl 直下の comments*.xml は openpyxl によって xl/comments/
        # フォルダ以下に移動されているので、復元しない。
        if cmt_ptn.fullmatch(part_name):
            continue
        if part_name not in after_overrides:
            after_root.append(child)

    # Add missing <Default> tags from original to modified
    for child in before_root.findall(".//{*}Default"):
        extension = child.get("Extension")
        if extension and extension not in after_defaults:
            after_root.append(child)

    # 保存する。
    tree = etree.ElementTree(after_root)
    tree.write(after_dir / '[Content_Types].xml', encoding='utf-8')


def restore_folder(
        before_dir: Path, after_dir: Path, folder2restore: str,
        delete_first=False):
    '''folder2restore 引数で指定されたフォルダを復元する。
    '''
    src = before_dir / folder2restore
    dest = after_dir / folder2restore

    if delete_first:
        shutil.rmtree(dest)

    if not os.path.exists(dest):
        shutil.copytree(src, dest)


def restore_workbook_xml_rels(
        before_dir: Path, after_dir: Path, filename: str):
    '''xl/_rels/workbook.xml.rels 内の要素を復元する。

    filename は sharedStrings.xml, metadata.xml のいずれか。
    xl/ フォルダの直下にそのファイルがあったが保存後になくなった場合:
        workbook.xml.rels に Relationship タグを追加。
        必要に応じて xl/ フォルダにこれらのファイルをコピー。
        復元した workbook.xml.rels を xl/_rels/ に保存。
    '''
    after_tree = etree.parse(after_dir / 'xl/_rels/workbook.xml.rels')
    after_root = after_tree.getroot()

    before_rel_path = before_dir / 'xl' / filename
    after_rel_path = after_dir / 'xl' / filename

    # ファイルが保存後に存在しているか、保存前に存在していないなら、中断。
    if after_rel_path.exists() or not before_rel_path.exists():
        return

    shutil.copy(before_rel_path, after_rel_path)

    # 保存後の workbook.xml.rels に Relationship があれば、中断。
    found = after_root.xpath(f'.//Relationship[@Target="{filename}"]')
    if found:
        return

    # Relationship 要素の新しい Id を採番する。
    ids = [
        int(re.search(r'\d+', rel.get("Id", ""))[0])  # 数値部分を抽出
        for rel in after_root.xpath(".//Relationship")
        if re.search(r'\d+', rel.get("Id", ""))
    ]
    max_id = max(ids) if ids else 0
    new_id = f'rId{max_id + 1}'

    NAMESPACE = \
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    type_ = 'sheetMetadata' if filename == 'metadata.xml' else filename[:-4]

    new_rel = etree.Element("Relationship", {
        "Id": new_id,
        "Type": f'{NAMESPACE}/{type_}',
        "Target": filename,
    })
    after_root.append(new_rel)

    # 保存する。
    tree = etree.ElementTree(after_root)
    tree.write(after_dir / 'xl/_rels/workbook.xml.rels', encoding='utf-8')


def restore_worksheets(before_dir: Path, after_dir: Path):
    '''xl/worksheets/ フォルダ下の sheet*.xml の drawing 要素を復元する。
    '''
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
        wb: Workbook, src: Path, dest: Path, temp_dir_args=None):
    '''図形や画像を復元しつつ Workbook を保存する。

    Parameters
    ----------
    wb : Workbook
        保存する Workbook。
    src : Path
        復元する図形や画像の元となるブックファイルのパス。
    dest : Path
        Workbook の保存先となるブックファイルのパス。
    temp_dir_args : dict | None, default None
        TemporaryDirectory を作る時のパラメータ。
    '''
    if temp_dir_args is None:
        temp_dir_args = {}

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
        restore_content_types(before_dir, after_dir)

        # xl/diagrams/ フォルダを復元 (before ⇒ after) する。
        restore_folder(before_dir, after_dir, 'xl/diagrams/')

        # xl/media/ フォルダを復元 (before ⇒ after) する。
        restore_folder(before_dir, after_dir, 'xl/media/')

        # xl/drawings/ フォルダを復元 (before ⇒ after) する。
        # 保存後にフォルダが存在するため、削除してから復元する。
        # TODO: 「削除してから復元」で矛盾が起きないか確認すること。
        restore_folder(
            before_dir, after_dir, 'xl/drawings/', delete_first=True)

        # xl/_rels/workbook.xml.rels の sharedStrings.xml を復元する。
        restore_workbook_xml_rels(before_dir, after_dir, 'sharedStrings.xml')

        # xl/_rels/metadata.xml.rels の metadata.xml を復元する。
        restore_workbook_xml_rels(before_dir, after_dir, 'metadata.xml')

        ''' **** ここまで動作確認済み。****
        '''

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
