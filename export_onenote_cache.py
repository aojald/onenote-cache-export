#!/usr/bin/env python3
"""Export the local Microsoft OneNote for Mac cache to plain-text files.

OneNote for Mac keeps a searchable copy of each notebook in a SQLite database
under ~/Library. This script reads that cache and writes one .txt file per
page, laid out as:

    <output>/<notebook>/<section>/<page>.txt

Each file's created and modified dates are set to the page date stored in the
cache. That date is the last time OneNote modified the page. The cache does
not keep a separate creation date, images, ink, or attachments.

Requires Python 3.9 or newer and only the standard library. Creation dates are
applied on macOS. On other systems only the modified date is set.

Examples:
    python3 export_onenote_cache.py --dry-run
    python3 export_onenote_cache.py --out ~/Desktop/OneNote-export
    python3 export_onenote_cache.py --db /path/to/cache.db --out ./notes --force
"""

from __future__ import annotations

import argparse
import ctypes
import os
import re
import sqlite3
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence
from urllib.parse import quote

PAGE = 1
SECTION = 2
NOTEBOOK = 4

# Windows FILETIME epoch (1601-01-01) to Unix epoch (1970-01-01), in seconds.
FILETIME_UNIX_OFFSET = 11_644_473_600
EARLIEST_UNIX = 631_152_000  # 1990-01-01
LATEST_UNIX = 4_102_444_800  # 2100-01-01

CACHE_ROOT = (
    Path.home()
    / "Library/Containers/com.microsoft.onenote.mac/Data/Library"
    / "Application Support/Microsoft User Data/OneNote"
)

WINDOWS_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

ILLEGAL_CHARS = re.compile(r'[\x00-\x1f\\/:*?"<>|]')


@dataclass(frozen=True)
class Page:
    rowid: int
    title: str
    notebook: str
    section: str
    modified: int | None
    paragraphs: tuple[str, ...]


def sqlite_uri(path: Path) -> str:
    encoded = quote(path.resolve().as_posix(), safe="/")
    return f"file:{encoded}?mode=ro"


def connect(path: Path) -> sqlite3.Connection:
    try:
        return sqlite3.connect(sqlite_uri(path), uri=True)
    except sqlite3.OperationalError:
        encoded = quote(path.resolve().as_posix(), safe="/")
        return sqlite3.connect(f"file:{encoded}?mode=ro&immutable=1", uri=True)


def discover_databases(root: Path = CACHE_ROOT) -> list[Path]:
    if not root.is_dir():
        return []
    found = []
    for path in root.glob("*/FullTextSearchIndex/*.db"):
        if path.is_file() and path.stat().st_size > 0 and not path.name.endswith("-journal"):
            found.append(path)
    return sorted(found)


def page_count(path: Path) -> int:
    try:
        with connect(path) as con:
            row = con.execute(
                "SELECT count(*) FROM Entities WHERE Type = ?", (PAGE,)
            ).fetchone()
        return int(row[0]) if row else -1
    except sqlite3.Error:
        return -1


def choose_database(candidates: Sequence[Path], explicit: Path | None) -> Path:
    if explicit is not None:
        if not explicit.is_file():
            raise SystemExit(f"Database not found: {explicit}")
        return explicit
    if not candidates:
        raise SystemExit(
            "No OneNote cache database found.\n"
            f"Looked in: {CACHE_ROOT}\n"
            "Pass --db if the cache lives somewhere else."
        )
    ranked = sorted(
        candidates,
        key=lambda path: (page_count(path), path.stat().st_mtime),
        reverse=True,
    )
    chosen = ranked[0]
    if len(ranked) > 1:
        print(f"Found {len(ranked)} cache databases. Using the one with the most pages:", file=sys.stderr)
        for path in ranked:
            marker = " (selected)" if path == chosen else ""
            print(f"  {page_count(path):5} pages  {path}{marker}", file=sys.stderr)
    return chosen


def require_schema(con: sqlite3.Connection) -> None:
    names = {
        row[0]
        for row in con.execute("SELECT name FROM sqlite_master WHERE type = 'table'")
    }
    missing = {"Entities", "PageElements"} - names
    if missing:
        raise SystemExit(
            "This database is not an OneNote full-text cache. "
            f"Missing table(s): {', '.join(sorted(missing))}"
        )
    entity_cols = {row[1] for row in con.execute("PRAGMA table_info(Entities)")}
    element_cols = {row[1] for row in con.execute("PRAGMA table_info(PageElements)")}
    needed_entities = {"Type", "GOID", "ParentGOID", "Title", "LastModifiedTime"}
    needed_elements = {"EntityRowId", "Text"}
    if not needed_entities <= entity_cols or not needed_elements <= element_cols:
        raise SystemExit("The cache schema is not the one this tool understands.")


def filetime_to_unix(value: int | None) -> int | None:
    if not value or value < 0:
        return None
    unix = int(value / 10_000_000) - FILETIME_UNIX_OFFSET
    if unix < EARLIEST_UNIX or unix > LATEST_UNIX:
        return None
    return unix


def safe_component(name: str, fallback: str, limit: int = 120) -> str:
    text = unicodedata.normalize("NFC", name or "")
    text = ILLEGAL_CHARS.sub("-", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    text = text.lstrip(".")
    if not text:
        text = fallback
    if text.upper() in WINDOWS_RESERVED:
        text = f"{text}-note"
    if len(text) > limit:
        text = text[:limit].rstrip(" .")
    return text or fallback


def load_pages(con: sqlite3.Connection) -> list[Page]:
    notebooks = {
        goid: title or ""
        for goid, title in con.execute(
            "SELECT GOID, Title FROM Entities WHERE Type = ?", (NOTEBOOK,)
        )
    }
    sections = {
        goid: (title or "", parent)
        for goid, title, parent in con.execute(
            "SELECT GOID, Title, ParentGOID FROM Entities WHERE Type = ?", (SECTION,)
        )
    }
    paragraphs: dict[int, list[str]] = {}
    for entity_id, text in con.execute(
        "SELECT EntityRowId, Text FROM PageElements "
        "WHERE Text IS NOT NULL AND length(Text) > 0 "
        "ORDER BY EntityRowId, rowid"
    ):
        cleaned = str(text).replace("\x00", "").strip()
        if cleaned and entity_id is not None:
            paragraphs.setdefault(int(entity_id), []).append(cleaned)

    pages: list[Page] = []
    for rowid, title, parent, modified in con.execute(
        "SELECT rowid, Title, ParentGOID, LastModifiedTime FROM Entities "
        "WHERE Type = ? ORDER BY ParentGOID, rowid",
        (PAGE,),
    ):
        section_title, section_parent = sections.get(parent, ("", None))
        notebook_title = notebooks.get(section_parent, "")
        pages.append(
            Page(
                rowid=int(rowid),
                title=(title or "").strip(),
                notebook=notebook_title.strip() or "Notebook",
                section=section_title.strip() or "Section",
                modified=filetime_to_unix(modified),
                paragraphs=tuple(paragraphs.get(int(rowid), ())),
            )
        )
    return pages


class BirthTime:
    """Set a file's creation date. macOS stores it separately from mtime."""

    def __init__(self) -> None:
        self._set = self._load()
        self._warned = False

    def _load(self):
        if sys.platform != "darwin":
            return None
        try:
            libc = ctypes.CDLL("/usr/lib/libSystem.B.dylib", use_errno=True)
        except OSError:
            return None

        class AttrList(ctypes.Structure):
            _fields_ = [
                ("bitmapcount", ctypes.c_ushort),
                ("reserved", ctypes.c_ushort),
                ("commonattr", ctypes.c_uint),
                ("volattr", ctypes.c_uint),
                ("dirattr", ctypes.c_uint),
                ("fileattr", ctypes.c_uint),
                ("forkattr", ctypes.c_uint),
            ]

        class TimeSpec(ctypes.Structure):
            _fields_ = [("tv_sec", ctypes.c_long), ("tv_nsec", ctypes.c_long)]

        attr_crtime = 0x00000200
        libc.setattrlist.argtypes = [
            ctypes.c_char_p,
            ctypes.POINTER(AttrList),
            ctypes.c_void_p,
            ctypes.c_size_t,
            ctypes.c_uint,
        ]
        libc.setattrlist.restype = ctypes.c_int

        def apply(path: Path, unix_ts: int) -> bool:
            attr = AttrList(bitmapcount=5, commonattr=attr_crtime)
            spec = TimeSpec(unix_ts, 0)
            rc = libc.setattrlist(
                os.fsencode(path),
                ctypes.byref(attr),
                ctypes.byref(spec),
                ctypes.sizeof(spec),
                0,
            )
            return rc == 0

        return apply

    def apply(self, path: Path, unix_ts: int) -> None:
        os.utime(path, (unix_ts, unix_ts))
        if self._set is None:
            return
        if self._set(path, unix_ts):
            return
        if not self._warned:
            print(
                "Could not set creation dates. Modified dates were still updated.",
                file=sys.stderr,
            )
            self._warned = True


class Paths:
    """Assign stable, collision-free folder and file names for one export."""

    def __init__(self) -> None:
        self._notebooks: dict[str, str] = {}
        self._sections: dict[tuple[str, str], str] = {}
        self._notebook_names: set[str] = set()
        self._section_names: dict[str, set[str]] = {}
        self._file_names: dict[tuple[str, str], set[str]] = {}

    def path(self, root: Path, page: Page) -> Path:
        notebook = self._notebooks.get(page.notebook)
        if notebook is None:
            notebook = allocate(safe_component(page.notebook, "Notebook"), self._notebook_names)
            self._notebooks[page.notebook] = notebook
            self._notebook_names.add(notebook.lower())

        section_key = (page.notebook, page.section)
        section = self._sections.get(section_key)
        if section is None:
            taken = self._section_names.setdefault(notebook, set())
            section = allocate(safe_component(page.section, "Section"), taken)
            self._sections[section_key] = section
            taken.add(section.lower())

        taken_files = self._file_names.setdefault((notebook, section), set())
        stem = allocate(safe_component(page.title, f"Untitled-{page.rowid}"), taken_files)
        taken_files.add(stem.lower())
        return root / notebook / section / f"{stem}.txt"


def allocate(stem: str, used: set[str]) -> str:
    if stem.lower() not in used:
        return stem
    number = 2
    while f"{stem}-{number}".lower() in used:
        number += 1
    return f"{stem}-{number}"


def render(page: Page) -> str:
    title = page.title or "Untitled"
    body = "\n\n".join(page.paragraphs)
    if body:
        return f"{title}\n\n{body}\n"
    return f"{title}\n"


def ensure_output(path: Path, force: bool, dry_run: bool) -> None:
    if path.exists() and not path.is_dir():
        raise SystemExit(f"Output path exists and is not a directory: {path}")
    if path.exists() and any(path.iterdir()) and not force and not dry_run:
        raise SystemExit(
            f"{path} is not empty.\n"
            "Pass --force to write into it, or choose another folder with --out."
        )
    if not dry_run:
        path.mkdir(parents=True, exist_ok=True)


def export(pages: Iterable[Page], root: Path, dry_run: bool) -> dict[str, int]:
    birth = BirthTime()
    paths = Paths()
    stats = {"written": 0, "empty": 0, "dated": 0, "undated": 0}
    for page in pages:
        path = paths.path(root, page)
        if not page.paragraphs:
            stats["empty"] += 1
        if page.modified is None:
            stats["undated"] += 1
        else:
            stats["dated"] += 1
        stats["written"] += 1
        if dry_run:
            continue
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(render(page), encoding="utf-8", newline="\n")
        if page.modified is not None:
            birth.apply(path, page.modified)
    return stats


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export the OneNote for Mac cache to a folder of .txt files.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "examples:\n"
            "  %(prog)s --dry-run\n"
            "  %(prog)s --out ~/Desktop/OneNote-export\n"
            "  %(prog)s --db ~/path/to/cache.db --out ./notes --force\n"
        ),
    )
    parser.add_argument(
        "--db",
        type=Path,
        help="OneNote full-text cache database. Found automatically when omitted.",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=Path.home() / "Desktop/OneNote-export",
        help="output folder (default: ~/Desktop/OneNote-export)",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="write into an output folder that already contains files",
    )
    parser.add_argument(
        "--skip-empty",
        action="store_true",
        help="skip pages that have a title but no text",
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="export at most this many pages",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="show what would be written without creating files",
    )
    args = parser.parse_args(argv)
    if args.limit is not None and args.limit < 1:
        parser.error("--limit must be at least 1")
    return args


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    database = choose_database(discover_databases(), args.db)
    with connect(database) as con:
        require_schema(con)
        pages = load_pages(con)
    if args.skip_empty:
        pages = [page for page in pages if page.paragraphs]
    if args.limit is not None:
        pages = pages[: args.limit]
    if not pages:
        print("The cache has no pages to export.")
        return 0

    output = args.out.expanduser()
    ensure_output(output, args.force, args.dry_run)
    stats = export(pages, output, args.dry_run)

    notebooks = sorted({page.notebook for page in pages})
    sections = sorted({(page.notebook, page.section) for page in pages})
    action = "Would write" if args.dry_run else "Wrote"
    print(f"Cache: {database}")
    print(f"{action} {stats['written']} pages in {len(sections)} sections, {len(notebooks)} notebook(s)")
    print(f"  with text: {stats['written'] - stats['empty']}")
    print(f"  title only: {stats['empty']}")
    print(f"  dated: {stats['dated']}")
    print(f"  no date in cache: {stats['undated']}")
    print(f"Output: {output}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nStopped.", file=sys.stderr)
        raise SystemExit(130)
    except sqlite3.Error as exc:
        print(f"Could not read the OneNote cache: {exc}", file=sys.stderr)
        raise SystemExit(1)
