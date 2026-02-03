# MD5 Tool (PyQt6)

Desktop utility that lets you pick files or folders, then hashes every file concurrently and displays MD5 results in a sortable table.

## Setup

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

## Usage

- Add files or folders (folder selection recurses into subdirectories).
- Click Start to hash; Cancel stops any remaining work.
- Click column headers to sort by path, size, hash, or duration.
- Worker count is auto-chosen based on CPU cores (2 to 32 threads).

## Notes

- Errors (permission issues, locks, cancellations) show in the Status column.
- Hashing stops early if Cancel is pressed; already completed rows stay filled.
- The app reads files in 128 KB chunks to balance throughput and memory.
- Logs are written to logs/md5tool.log (rotating up to ~2MB x 3).
