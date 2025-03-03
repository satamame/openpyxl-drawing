import copy
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree
from lxml.etree import Element
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
        part_name: str = child.get("PartName")
        if not part_name:
            continue
        # xl/calcChain.xml は Excel が復元するので、復元しない。
        if part_name == '/xl/calcChain.xml':
            continue
        # xl/ctrlProps/* は、ActiveX を使っていない想定なので復元しない。
        if part_name.startswith('/xl/ctrlProps/'):
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


def restore_xl_drawings_folder(before_dir: Path, after_dir: Path):
    '''xl/drawings/ フォルダを復元する。
    '''
    src = before_dir / 'xl/drawings/'
    dest = after_dir / 'xl/drawings/'

    shutil.rmtree(dest)
    dest.mkdir()

    for f in src.iterdir():
        # _rels/ フォルダと *.xml ファイルを復元する。
        if f.name == '_rels':
            shutil.copytree(f, dest / '_rels')
        elif f.suffix == '.xml':
            shutil.copy2(f, dest)


def get_rel_max_id(el: Element) -> int:
    '''ある xml 要素の下で、Relationship 要素の最大 Id を取得する。
    '''
    ids = [
        int(re.search(r'\d+', rel.get("Id", ""))[0])  # 数値部分を抽出
        for rel in el.xpath(".//Relationship")
        if re.search(r'\d+', rel.get("Id", ""))
    ]
    max_id = max(ids) if ids else 0
    return max_id


def adjust_worksheets(after_dir: Path):
    '''xl/worksheets/ フォルダ下の sheet*.xml の内容を調整する。
    '''
    dest_dir = after_dir / 'xl/worksheets'

    fname_ptn = re.compile(r'sheet[0-9]+\.xml')

    for f in dest_dir.iterdir():
        if not (f.is_file() and fname_ptn.fullmatch(f.name)):
            continue

        tree = etree.parse(f)
        root = tree.getroot()

        namespaces = {'ns': root.nsmap[None]}

        # まず、`<drawing>` 要素と `<legacyDrawing>` 要素を削除する。
        drawings = root.xpath("ns:drawing", namespaces=namespaces)
        for el in drawings:
            root.remove(el)
        legacy_drawings = root.xpath("ns:legacyDrawing", namespaces=namespaces)
        for el in legacy_drawings:
            root.remove(el)

        # 対応する sheet*.xml.rels ファイル
        rel_f = after_dir / f'xl/worksheets/_rels/{f.name}.rels'
        if not rel_f.is_file():
            continue

        rel_tree = etree.parse(rel_f)
        rel_root = rel_tree.getroot()

        # Target="../drawings/drawing*.xml" である Relationship を取得。
        namespaces = {'ns': rel_root.nsmap[None]}
        rels = rel_root.xpath('ns:Relationship', namespaces=namespaces)
        target_ptn = re.compile(r'\.\./drawings/drawing[0-9]+.xml')
        drawings = []
        for rel in rels:
            target = rel.get('Target')
            if target and target_ptn.fullmatch(target):
                drawings.append(rel)

        # sheet*.xml の root に namespace:r を追加して、新しい root を作る。
        rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/" + \
            "relationships"
        nsmap = root.nsmap.copy() if root.nsmap else {}
        nsmap["r"] = rel_ns

        # 新しい root を作成し直す（既存の要素を移植）
        new_root = etree.Element(root.tag, nsmap=nsmap)
        new_root.extend(root)

        # 新しい root に Relation と同じ id の drawing 要素を追加する。
        for drawing in drawings:
            el = etree.Element('drawing')
            el.set(f'{{{rel_ns}}}id', drawing.get('Id'))
            new_root.append(el)

        # 保存する。
        tree = etree.ElementTree(new_root)
        tree.write(f, encoding='utf-8')


def restore_sheet_xml_rels(before_dir: Path, after_dir: Path):
    '''xl/worksheets/_rels/sheet*.xml.rels 内の Relation を復元する。

    Target="../drawings/drawing*.xml" のものを復元する。
    '''
    src_dir = before_dir / 'xl/worksheets/_rels'
    dest_dir = after_dir / 'xl/worksheets/_rels'

    fname_ptn = re.compile(r'sheet[0-9]+\.xml.rels')

    for f in src_dir.iterdir():
        # sheet*.xml.rels ファイルだけを処理する。
        if not (f.is_file() and fname_ptn.fullmatch(f.name)):
            continue

        before_tree = etree.parse(f)
        before_root = before_tree.getroot()

        f2 = dest_dir / f.name
        if not f2.is_file():
            # 保存後にそのファイルがなければ新規で作る。
            after_tree = copy.deepcopy(before_tree)
            after_root = after_tree.getroot()
            after_root.clear()
        else:
            after_tree = etree.parse(f2)
            after_root = after_tree.getroot()

        # Target="../drawings/vmlDrawing*.vml" である Relationship を削除。
        # Target="/xl/drawings/vmlDrawing*.vml" になっている場合も考慮する。
        namespaces = {'ns': after_root.nsmap[None]}
        rels = after_root.xpath('ns:Relationship', namespaces=namespaces)
        target_ptn = re.compile(r'(\.\.|/xl)/drawings/vmlDrawing[0-9]+.vml')
        for rel in rels:
            target = rel.get("Target")
            if target and target_ptn.fullmatch(target):
                after_root.remove(rel)

        # 保存前の Target="../drawings/drawing*.xml" である Relationship を取得。
        namespaces = {'ns': before_root.nsmap[None]}
        rels = before_root.xpath('ns:Relationship', namespaces=namespaces)
        target_ptn = re.compile(r'\.\./drawings/drawing[0-9]+.xml')
        existings = []
        for rel in rels:
            target = rel.get('Target')
            if target and target_ptn.fullmatch(target):
                existings.append(rel)

        # 保存後の xml に取得した Relationship を足していく。
        # id は採番しなおす。
        max_id = get_rel_max_id(after_root)
        for rel in existings:
            # 保存後の sheet*.xml.rels に同じ Relationship があれば、スキップ。
            target = rel.get('Target')
            found = after_root.xpath(
                f'ns:Relationship[@Target="{target}"]', namespaces=namespaces)
            if found:
                continue

            max_id += 1
            rel.set('Id', f'rId{max_id}')
            after_root.append(rel)

        # 保存する。
        if len(after_root):
            tree = etree.ElementTree(after_root)
            tree.write(f2, encoding='utf-8')
        else:
            # root に子要素がなければファイルを削除する。
            f2.unlink(missing_ok=True)


def restore_doc_props_app(before_dir: Path, after_dir: Path):
    '''docProps/app.xml の重要な要素を復元する。
    '''
    before_app_path = before_dir / 'docProps/app.xml'
    after_app_path = after_dir / 'docProps/app.xml'

    before_tree = etree.parse(before_app_path)
    before_root = before_tree.getroot()

    after_tree = etree.parse(after_app_path)
    after_root = after_tree.getroot()

    # vt ネームスペースを取得
    vt_namespace = before_root.nsmap.get("vt")
    namespaces = after_root.nsmap
    if vt_namespace:
        namespaces['vt'] = vt_namespace

    # HeadingPairs と TitlesOfParts を取得
    heading_pairs = before_root.find(
        "ns:HeadingPairs", namespaces={"ns": before_root.nsmap[None]})
    titles_of_parts = before_root.find(
        "ns:TitlesOfParts", namespaces={"ns": before_root.nsmap[None]})

    # 保存後の root から HeadingPairs と TitlesOfParts を削除
    elems = after_root.findall(
        "ns:HeadingPairs", namespaces={"ns": after_root.nsmap[None]})
    for elem in elems:
        after_root.remove(elem)
    elems = after_root.findall(
        "ns:TitlesOfParts", namespaces={"ns": after_root.nsmap[None]})
    for elem in elems:
        after_root.remove(elem)

    # 保存前から取得した HeadingPairs と TitlesOfParts を保存後の root に追加
    if heading_pairs is not None:
        after_root.append(heading_pairs)

    if titles_of_parts is not None:
        after_root.append(titles_of_parts)

    # 保存する。
    tree = etree.ElementTree(after_root)
    tree.write(after_app_path, encoding='utf-8')


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
        restore_xl_drawings_folder(before_dir, after_dir)

        # xl/worksheets/_rels/sheet*.xml.rels 内の Relation を復元する。
        restore_sheet_xml_rels(before_dir, after_dir)

        # docProps/app.xml の重要な要素を復元する。
        restore_doc_props_app(before_dir, after_dir)

        # xl/ctrlProps/ フォルダを削除する
        # ※ ActiveX control を使っていない前提。
        shutil.rmtree(after_dir / 'xl/ctrlProps', ignore_errors=True)

        # xl/worksheets/sheet*.xml の内容を調整する。
        adjust_worksheets(after_dir)

        # dest に圧縮しなおす。
        with zipfile.ZipFile(dest, 'w') as zf:
            for root, _, files in os.walk(after_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, after_dir)
                    zf.write(file_path, arcname)
