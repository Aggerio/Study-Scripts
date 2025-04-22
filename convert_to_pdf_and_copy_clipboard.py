#!/usr/bin/env python3
"""
convert_and_extract.py

1. Converts all .pptx, .ppt and .docx in the specified folder to PDF (using libreoffice).
2. Moves the originals into subfolders original_pptx, original_ppt, original_docx.
3. Extracts text from each PDF into a .txt alongside it.
4. Concatenates all .txt files and copies to clipboard via xclip.
"""

import os
import sys
import argparse
import subprocess
from pathlib import Path
from PyPDF2 import PdfReader


def run_conversion(input_folder: Path, ext: str):
    """Convert all files with given extension to PDF and move originals."""
    files = list(input_folder.glob(f"*{ext}"))
    if not files:
        return
    # run libreoffice conversion
    cmd = ["libreoffice", "--headless", "--convert-to", "pdf", f'*.{ext.lstrip(".")}']
    print(f"Converting {ext} → PDF...")
    # Must run in the folder
    subprocess.run(" ".join(cmd), cwd=str(input_folder), shell=True, check=True)
    # move originals
    orig_dir = input_folder / f'original_{ext.lstrip(".")}'
    orig_dir.mkdir(exist_ok=True)
    for f in files:
        print(f"Moving original {f.name} → {orig_dir.name}/")
        f.rename(orig_dir / f.name)


def extract_text_from_pdf(pdf_path: Path) -> str:
    """Extract all text from a single PDF."""
    reader = PdfReader(str(pdf_path))
    chunks = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            chunks.append(text)
    return "\n".join(chunks)


def extract_all_pdfs(input_folder: Path):
    """For each PDF in folder, write a .txt of its text."""
    for pdf in input_folder.glob("*.pdf"):
        txt_path = pdf.with_suffix(".txt")
        print(f"Extracting text from {pdf.name} → {txt_path.name}")
        text = extract_text_from_pdf(pdf)
        txt_path.write_text(text, encoding="utf-8")


def copy_txts_to_clipboard(input_folder: Path):
    """Concatenate all .txt files and pipe into xclip."""
    txts = sorted(input_folder.glob("*.txt"))
    if not txts:
        print("No .txt files to copy.")
        return
    print(f"Copying {len(txts)} .txt files into clipboard...")
    # build one big string
    combined = []
    for txt in txts:
        combined.append(txt.read_text(encoding="utf-8"))
    data = "\n\n".join(combined)
    # call xclip
    p = subprocess.Popen(["xclip", "-selection", "clipboard"], stdin=subprocess.PIPE)
    p.communicate(input=data.encode("utf-8"))
    if p.returncode == 0:
        print("✅ All text copied to clipboard.")
    else:
        print("⚠️  Failed to copy to clipboard (xclip error).")


def main():
    p = argparse.ArgumentParser(
        description="Convert pptx/ppt/docx → PDF, extract text, copy to clipboard."
    )
    p.add_argument(
        "folder",
        nargs="?",
        default=os.getcwd(),
        help="Target folder (default: current directory)",
    )
    args = p.parse_args()
    folder = Path(args.folder).expanduser().resolve()
    if not folder.is_dir():
        print(f"Error: {folder} is not a directory.")
        sys.exit(1)

    # 1. conversion & moving originals
    for ext in [".pptx", ".ppt", ".docx"]:
        run_conversion(folder, ext)

    # 2. extract PDF → TXT
    extract_all_pdfs(folder)

    # 3. combine and copy to clipboard
    copy_txts_to_clipboard(folder)


if __name__ == "__main__":
    main()
