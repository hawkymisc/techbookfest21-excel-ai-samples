"""生成したファイルを一時フォルダで作り、検査を通ってから samples/ に移す。

Excel の保存・個人情報の除去・検査のどこかで失敗しても、作りかけのファイルが samples/ に残らないようにする
（前回作ったファイルもそのまま残る）。使い方：

    with staging.staging_dir() as stage:
        path = stage / "売上.xlsx"
        ...  # Excel で path に保存し、scrub_personal_info・verify_no_personal_info を通す
        staging.publish(path, SAMPLES / "売上.xlsx")

移動は os.replace（同じドライブなら置き換えが一度で済む）。一時フォルダが別ドライブにあって os.replace が
使えないときは、移動先のフォルダの隠しファイル（.<名前>.<乱数>.tmp。.gitignore で無視する）に shutil.move で
移してから os.replace する。
"""
import errno
import os
import shutil
import tempfile
from collections.abc import Iterator, Mapping
from contextlib import contextmanager
from pathlib import Path

PREFIX = "excel-ai-book-"
_ERROR_NOT_SAME_DEVICE = 17  # Windows の ERROR_NOT_SAME_DEVICE


@contextmanager
def staging_dir() -> Iterator[Path]:
    """一時フォルダを作って渡す。抜けるときに中身ごと消す（publish しなかったファイルは残らない）。"""
    with tempfile.TemporaryDirectory(prefix=PREFIX) as name:
        yield Path(name)


def _is_cross_device(error: OSError) -> bool:
    return error.errno == errno.EXDEV or getattr(error, "winerror", None) == _ERROR_NOT_SAME_DEVICE


def publish(src: Path, dest: Path) -> None:
    """一時フォルダのファイルを dest に移す（あれば置き換える）。移動先のフォルダがなければ作る。"""
    src, dest = Path(src), Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        os.replace(src, dest)
        return
    except OSError as e:
        if not _is_cross_device(e):
            raise
    fd, tmp_name = tempfile.mkstemp(dir=dest.parent, prefix=f".{dest.stem}.", suffix=".tmp")
    os.close(fd)
    tmp = Path(tmp_name)
    tmp.unlink()  # 名前だけを使う。空のファイルがあると Windows の rename が FileExistsError になる
    try:
        shutil.move(src, tmp)
        os.replace(tmp, dest)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise


def publish_all(moves: Mapping[Path, Path]) -> None:
    """{一時フォルダのファイル: 移動先} をすべて移す。全部の検査が済んでから呼ぶ。"""
    for src, dest in moves.items():
        publish(src, dest)
