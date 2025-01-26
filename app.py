import argparse
import tempfile
from datetime import datetime
from pathlib import Path

import openpyxl

from save_with_drawings import save_with_drawings


def main(src: Path, dest: Path, keep_temp_dir=False):
    wb = openpyxl.load_workbook(src)
    wb.worksheets[0]['A1'].value = datetime.now()

    delete = not keep_temp_dir

    with tempfile.TemporaryDirectory(
            prefix='temp_', dir='.', delete=delete) as temp_dir:
        save_with_drawings(wb, src, dest, Path(temp_dir))

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

    args = parser.parse_args()

    main(Path(args.src), Path(args.dest), args.keep)
