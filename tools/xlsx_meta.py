"""読者に配る xlsx から、Excel が保存時に書き込む個人情報を取り除き、残っていないことを確かめる。

取り除くもの：作成者（dc:creator）、最終更新者（cp:lastModifiedBy）、会社名（Company）、管理者（Manager）、
保存先の絶対パス（x15ac:absPath）。作成日時・更新日時は固定値に置き換え、zip の各エントリの日時もそろえる
（作り直したときにバイナリの差分が出にくくするため）。
"""
import os
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from collections.abc import Iterable
from pathlib import Path

FIXED_TIMESTAMP = "2026-10-01T00:00:00Z"
ZIP_DATE_TIME = (1980, 1, 1, 0, 0, 0)

NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
}

# 個人情報を書き換える要素（属性付き・自己終了形式も対象）
PERSON_ELEMENTS = re.compile(r"<(dc:creator|cp:lastModifiedBy)\b[^>]*?(?:/>|>[^<]*</\1\s*>)")
APP_ELEMENTS = re.compile(r"<(Company|Manager)\b[^>]*?(?:/>|>[^<]*</\1\s*>)")
TIMESTAMP_ELEMENTS = re.compile(r"<(dcterms:created|dcterms:modified)\b([^>]*?)(?:/>|>[^<]*</\1\s*>)")
ABS_PATH_BLOCK = re.compile(
    r"<mc:AlternateContent\b[^>]*>\s*<mc:Choice\b[^>]*>\s*<x15ac:absPath\b[^>]*/>\s*</mc:Choice\s*>\s*</mc:AlternateContent\s*>"
)
ABS_PATH_ELEMENT = re.compile(r"<x15ac:absPath\b[^>]*/>")

# 保存先パスらしい文字列。名前空間の URL（http://...）の「p:/」に当たらないよう、ドライブ文字の前に英字がないことを条件にする
PATH_PATTERN = re.compile(
    r"(?i)(?<![a-z])[a-z]:[\\/]|file:/|\\\\[^\\]+\\|/Users/|d\.docs\.live\.net|sharepoint\.com|absPath"
)
# 配るブックに入っていてはいけないパート（外部リンク・プリンター設定・コメントの作成者など）
UNEXPECTED_PART = re.compile(
    r"(^|/)(externalLinks|printerSettings|persons|threadedComments|customXml)/|(^|/)comments[^/]*$|^docProps/custom\.xml$"
)


def _scrub_part(name: str, data: bytes) -> bytes:
    if name == "docProps/core.xml":
        text = data.decode("utf-8")
        text = PERSON_ELEMENTS.sub(r"<\1></\1>", text)
        text = TIMESTAMP_ELEMENTS.sub(lambda m: f"<{m.group(1)}{m.group(2)}>{FIXED_TIMESTAMP}</{m.group(1)}>", text)
        return text.encode("utf-8")
    if name == "docProps/app.xml":
        return APP_ELEMENTS.sub(r"<\1></\1>", data.decode("utf-8")).encode("utf-8")
    if name == "xl/workbook.xml":
        text = ABS_PATH_BLOCK.sub("", data.decode("utf-8"))
        return ABS_PATH_ELEMENT.sub("", text).encode("utf-8")
    return data


def scrub_personal_info(path: Path) -> None:
    """個人情報を xlsx から取り除く。同じディレクトリの一時ファイルに書いてから置き換える。"""
    path = Path(path)
    with zipfile.ZipFile(path) as src:
        entries = [(info, src.read(info.filename)) for info in src.infolist()]
    fd, tmp_name = tempfile.mkstemp(dir=path.parent, prefix=f".{path.stem}.", suffix=".tmp")
    tmp = Path(tmp_name)
    try:
        with os.fdopen(fd, "wb") as f, zipfile.ZipFile(f, "w", zipfile.ZIP_DEFLATED) as dst:
            for info, data in entries:
                out = zipfile.ZipInfo(info.filename, date_time=ZIP_DATE_TIME)
                out.compress_type = zipfile.ZIP_DEFLATED
                out.external_attr = info.external_attr
                dst.writestr(out, _scrub_part(info.filename, data))
        os.replace(tmp, path)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise


def _element_texts(xml: bytes, tag: str) -> list[str]:
    root = ET.fromstring(xml)
    prefix, local = tag.split(":")
    return [(e.text or "").strip() for e in root.iter(f"{{{NS[prefix]}}}{local}")]


def verify_no_personal_info(path: Path, user_names: Iterable[str] | None = None) -> None:
    """個人情報が残っていれば RuntimeError を送出する。

    user_names を省略すると、実行しているユーザーの名前（ホームディレクトリ名）がどのパートにも含まれないことも確かめる。
    """
    path = Path(path)
    if user_names is None:
        user_names = (Path.home().name,)
    names_to_find = [n.lower() for n in user_names if len(n) >= 3]  # 短すぎる名前は誤検出が多いので見ない
    with zipfile.ZipFile(path) as z:
        parts = {name: z.read(name) for name in z.namelist()}

    unexpected = sorted(n for n in parts if UNEXPECTED_PART.search(n))
    if unexpected:
        raise RuntimeError(f"{path.name}: 想定外のパートがある: {', '.join(unexpected)}")

    core = parts.get("docProps/core.xml")
    if core is None:
        raise RuntimeError(f"{path.name}: docProps/core.xml がない")
    for tag in ("dc:creator", "cp:lastModifiedBy"):
        if any(_element_texts(core, tag)):
            raise RuntimeError(f"{path.name}: docProps/core.xml の {tag} が空になっていない")
    for tag in ("dcterms:created", "dcterms:modified"):
        values = _element_texts(core, tag)
        if any(v != FIXED_TIMESTAMP for v in values):
            raise RuntimeError(f"{path.name}: docProps/core.xml の {tag} が固定値 {FIXED_TIMESTAMP} になっていない: {values}")

    app = parts.get("docProps/app.xml")
    if app is not None:
        for tag in ("Company", "Manager"):
            if any(_element_texts(app, f"ep:{tag}")):
                raise RuntimeError(f"{path.name}: docProps/app.xml の {tag} が空になっていない")

    for name, data in parts.items():
        if not name.endswith((".xml", ".rels")):
            continue
        text = data.decode("utf-8", errors="replace")
        m = PATH_PATTERN.search(text)
        if m:
            raise RuntimeError(f"{path.name}: {name} に保存先パスが残っている（{m.group(0)!r}）")
        lowered = text.lower()
        for user in names_to_find:
            if user in lowered:
                raise RuntimeError(f"{path.name}: {name} に実行ユーザー名が含まれている")
