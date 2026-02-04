#!/usr/bin/env python3
"""Convert Word documents (.docx) to LaTeX format using pandoc."""

import argparse
import subprocess
import sys
import zipfile
from pathlib import Path


def convert_docx_to_latex(
    input_file: Path,
    output_dir: Path | None = None,
    standalone: bool = True,
    extract_media: bool = True,
) -> Path:
    """
    Convert a Word document to LaTeX format.

    Creates a subdirectory named after the document (minus .docx) containing
    the .tex file and a media/ subdirectory for extracted images.

    Args:
        input_file: Path to the .docx file
        output_dir: Parent directory for output (default: same as input file)
        standalone: Include LaTeX preamble for compilable document
        extract_media: Extract images from the document

    Returns:
        Path to the generated .tex file
    """
    input_path = Path(input_file)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if not input_path.suffix.lower() == ".docx":
        raise ValueError(f"Input file must be .docx format: {input_path}")

    # Create subdirectory named after the document
    doc_name = input_path.stem
    parent_dir = Path(output_dir) if output_dir else input_path.parent
    doc_dir = parent_dir / doc_name
    doc_dir.mkdir(parents=True, exist_ok=True)

    output_path = doc_dir / f"{doc_name}.tex"

    cmd = ["pandoc", str(input_path)]

    if standalone:
        cmd.append("-s")

    if extract_media:
        cmd.extend(["--extract-media", str(doc_dir)])

    cmd.extend(["-o", str(output_path)])

    print(f"Converting: {input_path.name} -> {doc_name}/{output_path.name}")

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        raise RuntimeError(f"Pandoc error: {result.stderr}")

    print(f"Successfully created: {output_path}")

    # Create zip archive
    zip_path = create_zip_archive(doc_dir, doc_name)
    print(f"Created archive: {zip_path}")

    return output_path


def create_zip_archive(doc_dir: Path, doc_name: str) -> Path:
    """
    Create a zip archive containing the .tex file and media directory.

    Args:
        doc_dir: Directory containing the converted files
        doc_name: Base name for the document

    Returns:
        Path to the created zip file
    """
    zip_path = doc_dir / f"{doc_name}.zip"

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        # Add the .tex file
        tex_file = doc_dir / f"{doc_name}.tex"
        if tex_file.exists():
            zf.write(tex_file, tex_file.name)

        # Add media directory contents
        media_dir = doc_dir / "media"
        if media_dir.exists():
            for media_file in media_dir.rglob("*"):
                if media_file.is_file():
                    arcname = f"media/{media_file.relative_to(media_dir)}"
                    zf.write(media_file, arcname)

    return zip_path


def batch_convert(
    input_dir: Path,
    output_dir: Path | None = None,
    **kwargs,
) -> list[Path]:
    """
    Convert all .docx files in a directory to LaTeX.

    Each document gets its own subdirectory containing the .tex file
    and a media/ subdirectory for extracted images.

    Args:
        input_dir: Directory containing .docx files
        output_dir: Parent directory for output subdirectories (default: same as input)
        **kwargs: Additional arguments passed to convert_docx_to_latex

    Returns:
        List of paths to generated .tex files
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir) if output_dir else input_path

    if not input_path.is_dir():
        raise NotADirectoryError(f"Not a directory: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)

    docx_files = list(input_path.glob("*.docx"))

    if not docx_files:
        print(f"No .docx files found in {input_path}")
        return []

    print(f"Found {len(docx_files)} Word document(s)\n")

    converted = []
    for docx_file in docx_files:
        try:
            result = convert_docx_to_latex(docx_file, output_path, **kwargs)
            converted.append(result)
        except Exception as e:
            print(f"Error converting {docx_file.name}: {e}", file=sys.stderr)

    return converted


def main():
    parser = argparse.ArgumentParser(
        description="Convert Word documents to LaTeX using pandoc"
    )
    parser.add_argument(
        "input",
        help="Input .docx file or directory containing .docx files",
    )
    parser.add_argument(
        "-o", "--output",
        help="Output directory (default: same location as input)",
    )
    parser.add_argument(
        "--no-standalone",
        action="store_true",
        help="Don't include LaTeX preamble (fragment only)",
    )
    parser.add_argument(
        "--no-media",
        action="store_true",
        help="Don't extract media files",
    )

    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else None

    kwargs = {
        "standalone": not args.no_standalone,
        "extract_media": not args.no_media,
    }

    try:
        if input_path.is_dir():
            results = batch_convert(input_path, output_path, **kwargs)
            print(f"\nConverted {len(results)} file(s)")
        else:
            convert_docx_to_latex(input_path, output_path, **kwargs)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
