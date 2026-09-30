# OneNote cache export

Export the local cache of **Microsoft OneNote for Mac** to plain-text files. One `.txt` file is written per page, in a folder per notebook and section. Each file’s created and modified dates are set from the date OneNote stored for that page.

The script uses only the Python standard library. It reads the cache on this Mac and writes files to a folder you choose. It does not upload notebook contents.

## What you get

```text
OneNote-export/
  Notebook name/
    Section name/
      Page title.txt
```

A page file is plain text. The first line is the page title, then a blank line, then the outline text. It is not Markdown.

## What the cache does not contain

OneNote’s full-text cache is a search index, not a full notebook backup.

- Images, ink, embedded files, and attachments are not exported.
- Layout, tables, and lists become separate paragraphs.
- The stored date is the last time the page was modified. The cache has no separate creation date.
- Notebooks that were never indexed, or that live only as `.one` files elsewhere, are not included.

## Requirements

- macOS, with OneNote for Mac installed and notebooks that have been opened on this computer
- Python 3.9 or newer

On other systems the script can still read a database you pass with `--db`. Only macOS gets the file creation date. The modified date is set everywhere.

## Install

Copy `export_onenote_cache.py` anywhere, or clone this repository and enter the `onenote-cache-export` directory.

No packages to install.

## Use

See what would be exported:

```bash
python3 export_onenote_cache.py --dry-run
```

Write the files:

```bash
python3 export_onenote_cache.py --out ~/Desktop/OneNote-export
```

The cache database is found automatically. It usually lives under:

```text
~/Library/Containers/com.microsoft.onenote.mac/Data/Library/Application Support/Microsoft User Data/OneNote/<version>/FullTextSearchIndex/*.db
```

If several caches exist, the one with the most pages is used. The others are listed so you can pass a specific file with `--db`.

### Options

| Option | Meaning |
|---|---|
| `--out DIR` | Output folder. Default: `~/Desktop/OneNote-export` |
| `--db FILE` | Cache database to read instead of the one found automatically |
| `--dry-run` | Print the counts and do not write files |
| `--force` | Write into an output folder that already contains files |
| `--skip-empty` | Skip pages that have a title and no text |
| `--limit N` | Export at most N pages |

The script refuses to write into a folder that already has files unless you pass `--force`.

### Examples

```bash
python3 export_onenote_cache.py --out ~/Documents/notes
python3 export_onenote_cache.py --db ~/path/to/cache.db --out ./notes
python3 export_onenote_cache.py --skip-empty --out ~/Desktop/OneNote-export
python3 export_onenote_cache.py --limit 20 --out /tmp/onenote-sample
```

## File names and dates

Notebook, section, and page titles are turned into folder and file names. Characters that are illegal in file names are replaced. Very long names are shortened. Names that differ only by letter case are kept distinct with a numeric suffix, because macOS disks are usually case-insensitive. Names reserved by Windows (`CON`, `AUX`, `NUL`, and similar) are adjusted so the folder can be copied to Windows.

A page whose title is empty is named `Untitled-<id>.txt`.

When the cache has a usable date, that instant becomes the file’s created time and modified time. Dates outside 1990–2100 are ignored. Pages with no date keep the time of the export.

## Privacy

Notebook text is written only to the output folder on this computer. Review that folder before you copy it anywhere else. This repository does not contain sample notebooks.

## Limits

This was checked against the OneNote for Mac full-text index (SQLite tables `Entities` and `PageElements`). A future OneNote release can change that schema. If the database shape does not match, the script stops instead of writing a partial export.

`.one` section files are a different, binary format. This tool does not read them.
