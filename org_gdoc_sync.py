import argparse
import os
import pypandoc
import re
import sys
from pathlib import Path

def get_document_title_and_id(path):
    """Generates a title and a CUSTOM_ID from a file path."""
    title = Path(path).stem
    # Create a simple, clean ID for org-mode
    custom_id = re.sub(r'[^a-zA-Z0-9]+', '-', title).lower().strip('-')
    return title, custom_id

def convert_docx_to_org(source_path, existing_org_file=None):
    """
    Converts a .docx file to a string of Org-mode content.
    It demotes headings (GDoc Heading 1 -> Org level 2) and uses the first
    heading as the basis for the top-level (level 1) entry.
    """
    abs_source_path = str(Path(source_path).resolve())
    title, custom_id = get_document_title_and_id(source_path)

    # Convert docx to org-mode, demoting headings by 1 level
    # (H1 -> **, H2 -> ***, etc.)
    # We add --wrap=none to prevent pandoc from hard-wrapping lines.
    org_content = pypandoc.convert_file(
        source_path, 'org', extra_args=["--base-header-level=2", "--wrap=none"]
    )

    content_lines = org_content.strip().split('\n')
    main_heading = f"* {title}"  # Fallback heading
    cleaned_content = org_content

    # Find the first demoted heading (**) to use as the chapter title.
    first_heading_text = None
    for i, line in enumerate(content_lines):
        match = re.match(r'^\*\*\s+(.*)', line)
        if match:
            first_heading_text = match.group(1).strip()
            # Remove this line as it's now promoted to the main heading
            content_lines.pop(i)
            cleaned_content = "\n".join(content_lines)
            break # Found the first one, stop looking

    if first_heading_text:
        main_heading = f"* {first_heading_text}"
    
    # Assemble the final, complete subtree content
    final_org_content = (
        f"{main_heading}\n"
        f":PROPERTIES:\n"
        f":SOURCE_FILE: {abs_source_path}\n"
        f":CUSTOM_ID: {custom_id}\n"
        f":END:\n\n"
        f"{cleaned_content.strip()}"
    )

    return final_org_content


def get_existing_sources(org_file_path):
    """Parses an org file and returns a set of existing SOURCE_FILE paths."""
    sources = set()
    if not os.path.exists(org_file_path):
        return sources
    try:
        with open(org_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            # Use regex to find all SOURCE_FILE properties
            found_paths = re.findall(r':SOURCE_FILE:\s+(.*)', content)
            for path in found_paths:
                # Resolve the path to handle tilde, etc.
                sources.add(str(Path(path.strip()).resolve()))
    except IOError as e:
        print(f"Warning: Could not read existing org file: {e}")
    return sources

def import_documents(args):
    """Main function to handle the import logic."""
    output_path = Path(args.output).expanduser().resolve()
    existing_sources = get_existing_sources(output_path)
    files_to_import = []

    if args.folder:
        folder_path = Path(args.folder).expanduser().resolve()
        if not folder_path.is_dir():
            print(f"Error: Folder not found at '{args.folder}'")
            return
        print(f"Importing from {folder_path}...")
        for docx_file in sorted(folder_path.glob('*.docx')):
            if str(docx_file.resolve()) not in existing_sources:
                files_to_import.append(docx_file)
    elif args.file:
        file_path = Path(args.file).expanduser().resolve()
        if not file_path.is_file():
            print(f"Error: File not found at '{args.file}'")
            return
        if str(file_path.resolve()) not in existing_sources:
            files_to_import.append(file_path)

    if not files_to_import:
        print("No new documents to import.")
        return

    # Open in append mode, creating if it doesn't exist
    with open(output_path, 'a', encoding='utf-8') as f:
        # If the file is new or empty, don't add initial newlines
        if f.tell() > 0:
            f.write("\n")
        for doc_path in files_to_import:
            print(f"  - Converting {doc_path.name}")
            org_content = convert_docx_to_org(doc_path)
            f.write(org_content + "\n")

    print(f"Successfully imported {len(files_to_import)} new document(s) into {output_path}")

def export_document(content, target_file, reference_doc):
    """Exports a region of org content to a .docx file."""
    # This logic removes the top-level heading and its properties drawer
    # so they don't appear as text in the exported docx file.
    lines = content.strip().split('\n')
    content_start_index = 0

    # Find the start of the actual content, skipping the Level 1 heading and properties
    if lines and lines[0].startswith('* '):
        content_start_index = 1
        in_drawer = False
        for i in range(1, len(lines)):
            line = lines[i].strip()
            if line.lower() == ':properties:':
                in_drawer = True
            elif line.lower() == ':end:' and in_drawer:
                content_start_index = i + 1
                break
            elif not in_drawer: # We are past the heading but not in a drawer
                content_start_index = i
                break

    cleaned_content = '\n'.join(lines[content_start_index:])

    extra_args = ["--wrap=none"]
    if reference_doc:
        extra_args.extend(["--reference-doc", str(Path(reference_doc).expanduser().resolve())])
    
    # We shift headings up so that a level 2 org heading (**) becomes Heading 1 in docx.
    extra_args.append("--shift-heading-level-by=-1")

    pypandoc.convert_text(
        cleaned_content,
        'docx',
        format='org',
        outputfile=str(Path(target_file).expanduser().resolve()),
        extra_args=extra_args
    )

def main():
    parser = argparse.ArgumentParser(description='Sync between Org-mode and a folder of .docx files.')
    subparsers = parser.add_subparsers(dest='command', required=True)

    # Import command
    parser_import = subparsers.add_parser('import', help='Import .docx files into an Org file.')
    group = parser_import.add_mutually_exclusive_group(required=True)
    group.add_argument('--folder', help='Path to the folder containing .docx files.')
    group.add_argument('--file', help='Path to a single .docx file to import.')
    parser_import.add_argument('--output', required=True, help='Path for the destination .org file.')

    # Export command
    parser_export = subparsers.add_parser('export', help='Export an Org subtree to a .docx file.')
    parser_export.add_argument('--target-file', required=True, help='The .docx file to write to.')
    parser_export.add_argument('--reference-doc', help='Path to a .docx style reference file.')
    parser_export.add_argument('--quiet', action='store_true', help='Suppress success messages.')

    args = parser.parse_args()

    if args.command == 'import':
        import_documents(args)
    elif args.command == 'export':
        # Read content from stdin for export using Python's standard library
        # for better compatibility with older pypandoc versions.
        org_content = sys.stdin.read()
        export_document(org_content, args.target_file, args.reference_doc)
        if not args.quiet:
            print(f"Successfully synced content to '{args.target_file}'.")

if __name__ == '__main__':
    main()

