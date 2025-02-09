import tempfile
import zipfile
from pathlib import Path

from openpyxl.workbook.workbook import Workbook


def save_with_openpyxl(
        wb: Workbook, src: Path, dest: Path, temp_dir_args=None):
    '''openpyxl で wb を save() する。

    Parameters
    ----------
    wb : Workbook
        保存する Workbook。
    src : Path
        元となるブックファイルのパス。
    dest : Path
        Workbook の保存先となるブックファイルのパス。
    temp_dir_args : dict | None, default None
        TemporaryDirectory を作る時のパラメータ。
    '''
    if temp_dir_args is None:
        temp_dir_args = {}

    with tempfile.TemporaryDirectory(**temp_dir_args) as temp_dir:
        src_dir = Path(temp_dir) / 'src'
        dest_dir = Path(temp_dir) / 'dest'

        # src を src_dir に解凍する。
        with zipfile.ZipFile(src, 'r') as zf:
            zf.extractall(str(src_dir))

        # wb を dest に保存する。
        wb.save(dest)

        # dest を dest_dir に解凍する。
        with zipfile.ZipFile(dest, 'r') as zf:
            zf.extractall(str(dest_dir))
