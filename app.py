import argparse
from datetime import datetime
from pathlib import Path

import openpyxl

from save_with_drawings import save_with_drawings
from save_with_openpyxl import save_with_openpyxl


def main(src: Path, dest: Path, keep_temp_dir=False, just_save=False):
    wb = openpyxl.load_workbook(
        src, keep_vba=True, rich_text=True, keep_links=True)

    temp_dir_args = {
        'prefix': 'temp_',
        'dir': '.',
        'delete': not keep_temp_dir,
    }

    if just_save:
        save_with_openpyxl(wb, src, dest, temp_dir_args)
    else:
        wb.worksheets[0]['A1'].value = datetime.now()
        save_with_drawings(wb, src, dest, temp_dir_args)

    wb.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        prog='Insert datetime',
        description='Excel シートの A1 セルに日時をセットする。')
    parser.add_argument('src', help='入力となる Excel ブックのファイル名。')
    parser.add_argument('dest', help='出力となる Excel ブックのファイル名。')
    parser.add_argument(
        '--keep-temp-dir', dest='keep', action='store_true',
        help='一時フォルダを削除しない。')
    parser.add_argument(
        '--just-save', dest='just', action='store_true',
        help='何もせず保存だけする。')

    args = parser.parse_args()

    main(
        Path(args.src), Path(args.dest),
        keep_temp_dir=args.keep, just_save=args.just)
