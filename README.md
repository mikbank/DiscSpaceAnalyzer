# Disc Space Analyser

Running low on disc space? Can't figure out what's eating up all your hard
drive space? Try running this script and get a simple report on the largest
folders on your computer.

Point it at a drive (or any folder), and it spits out an Excel sheet listing
every folder above a size threshold (1 GB by default), largest first. No more
manually right-clicking folders and waiting for "Properties" to calculate
size one at a time.

## Usage

1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
2. Copy `.env.example` to `.env` and adjust as needed:
   ```
   SCAN_ROOT=C:\
   SIZE_THRESHOLD_GB=1
   OUTPUT_FILE=large_folders_report.xlsx
   WORKERS=8
   ```
3. Run it:
   ```
   python main.py
   ```

It prints progress as it works, then writes `large_folders_report.xlsx`
(or whatever `OUTPUT_FILE` is set to) with three columns: folder path, size
in GB, and size in bytes, sorted largest first.

## How it works

For every folder in the tree, a folder's size is its own files plus the
(already-computed) size of every subfolder — computed bottom-up in one pass
using `os.scandir`, reading each file's size straight off the same directory
listing that found it, rather than doing a second, separate lookup per file
(`os.path.getsize`). On Windows this roughly halves the number of filesystem
calls, which matters a lot at scale: with antivirus intercepting most file
access, this is the main cost when a drive has hundreds of thousands of
folders.

### Why multiple worker processes

A single CPU core scanning a 200,000+ folder drive is slow mostly because
it's waiting on the filesystem/antivirus for each folder, one at a time.
Splitting the work across multiple processes (`WORKERS`, default 8) lets
those waits overlap instead of queuing up serially.

The tricky part is *how* to split the tree fairly. A drive's folders aren't
evenly spread out — most of a Windows install's size usually sits inside a
small number of folders (`C:\Windows`, `C:\Users\<you>`, `Program Files`,
...) while everything else is comparatively tiny. Splitting only at the top
level would leave most workers idle while one of them churns through a
folder that holds 90% of the data.

Instead, each worker tries to scan a folder fully, but bails out early if it
turns out to hold more than ~2,000 subfolders. When that happens, it reports
that folder's direct children back as new, smaller jobs instead of finishing
the scan itself. Those children get handed out to the pool the same way, so
a folder keeps getting split -- however many levels deep that takes -- until
every piece left is small enough to finish quickly. Children are handed out
in batches (100 at a time) rather than one job per child, since a single
folder can have tens of thousands of tiny direct children (Windows' `WinSxS`
component store is a real example) and creating a separate job per child
would itself become the bottleneck.

## Platform notes

This was written for and tested only on Windows. It has not been tested on
macOS or Linux -- the scanning logic itself is cross-platform (plain
`os.scandir`), but the symlink-skipping behavior and general performance
characteristics (antivirus overhead, etc.) were only verified on Windows.
