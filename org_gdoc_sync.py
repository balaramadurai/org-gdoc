import argparse
import os
import pypandoc
import re
import sys
import tempfile
from pathlib import Path

def get_document_title_and_id(path):
    """Generates a title and a CUSTOM_ID from a file path."""
    title = Path(path).stem
    # Create a simple, clean ID for org-mode
    custom_id = re.sub(r'[^a-zA-Z0-9]+', '-', title).lower().strip('-')
    return title, custom_id

def convert_docx_to_org(source_path, existing_org_file=None):
    """
    Converts a .docx file to a string of Org-mode content by reliably extracting
    metadata and body content in separate, robust steps.
    """
    abs_source_path = str(Path(source_path).resolve())
    fallback_title, custom_id = get_document_title_and_id(source_path)

    doc_title = None
    doc_subtitle = None

    # --- New, robust metadata extraction using a pandoc template ---
    # Create a temporary template to extract title and subtitle variables.
    metadata_template_content = "TITLE_VAR:$title$\nSUBTITLE_VAR:$subtitle$"
    with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix=".txt", encoding='utf-8') as temp_template:
        temp_template.write(metadata_template_content)
        metadata_template_path = temp_template.name

    try:
        # Extract metadata by running pandoc with the custom template via extra_args
        # for broader pypandoc version compatibility.
        metadata_output = pypandoc.convert_file(
            source_path, 'plain',
            extra_args=[f"--template={metadata_template_path}"]
        ).strip()
        for line in metadata_output.split('\n'):
            if line.startswith("TITLE_VAR:"):
                val = line[10:].strip()
                if val: doc_title = val
            elif line.startswith("SUBTITLE_VAR:"):
                val = line[13:].strip()
                if val: doc_subtitle = val
    finally:
        os.remove(metadata_template_path)
    # --- End of metadata extraction ---

    # --- New, clean body extraction using a pandoc template ---
    body_template_content = "$body$"
    with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix=".txt", encoding='utf-8') as temp_template:
        temp_template.write(body_template_content)
        body_template_path = temp_template.name

    try:
        # Extract just the body content, demoting headings.
        # Pass template via extra_args for compatibility.
        body_extra_args = [
            f"--template={body_template_path}",
            "--wrap=none",
            "--base-header-level=2"
        ]
        body_content = pypandoc.convert_file(
            source_path, 'org',
            extra_args=body_extra_args
        ).strip()
    finally:
        os.remove(body_template_path)
    # --- End of body extraction ---

    # Determine the main heading for the subtree.
    main_heading_title = doc_title if doc_title else fallback_title
    main_heading = f"* {main_heading_title}"

    properties = [
        f":SOURCE_FILE: {abs_source_path}",
        f":CUSTOM_ID: {custom_id}"
    ]
    if doc_title:
        properties.append(f":EXPORT_TITLE: {doc_title}")
    if doc_subtitle:
        properties.append(f":EXPORT_SUBTITLE: {doc_subtitle}")
    
    properties_drawer = ":PROPERTIES:\n" + "\n".join(properties) + "\n:END:"

    # Assemble the final, complete subtree content.
    final_org_content = f"{main_heading}\n{properties_drawer}"
        
    if body_content:
        final_org_content += f"\n\n{body_content}"

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
        files_to_import.append(file_path)

    if not files_to_import:
        if args.folder:
            print("No new documents to import.")
        return

    mode = 'a' if args.folder else 'w'
    with open(output_path, mode, encoding='utf-8') as f:
        if mode == 'a' and f.tell() > 0:
            f.write("\n")
        for doc_path in files_to_import:
            print(f"  - Converting {doc_path.name}")
            org_content = convert_docx_to_org(doc_path)
            f.write(org_content + "\n")

    if args.folder:
        print(f"Successfully imported {len(files_to_import)} new document(s) into {output_path}")

def export_document(content, target_file, reference_doc):
    """Exports a region of org content to a .docx file."""
    lines = content.strip().split('\n')
    
    metadata = {}
    content_lines = []
    
    in_drawer = False
    # Process lines, skipping the main heading (lines[0])
    for line in lines[1:]:
        stripped_line = line.strip()

        if stripped_line.lower() == ':properties:':
            in_drawer = True
            continue

        if stripped_line.lower() == ':end:' and in_drawer:
            in_drawer = False
            continue

        if in_drawer:
            title_match = re.match(r':EXPORT_TITLE:\s*(.*)', stripped_line, re.IGNORECASE)
            subtitle_match = re.match(r':EXPORT_SUBTITLE:\s*(.*)', stripped_line, re.IGNORECASE)
            if title_match:
                metadata['title'] = title_match.group(1).strip()
            elif subtitle_match:
                metadata['subtitle'] = subtitle_match.group(1).strip()
        else:
            content_lines.append(line)
            
    cleaned_content = '\n'.join(content_lines)

    extra_args = ["--wrap=none"]
    if reference_doc:
        extra_args.extend(["--reference-doc", str(Path(reference_doc).expanduser().resolve())])

    if 'title' in metadata:
        extra_args.extend(["-M", f"title={metadata['title']}"])
    if 'subtitle' in metadata:
        extra_args.extend(["-M", f"subtitle={metadata['subtitle']}"])

    extra_args.append("--shift-heading-level-by=-1")

    pypandoc.convert_text(
        cleaned_content,
        'docx',
        format='org',
        outputfile=str(Path(target_file).expanduser().resolve()),
        extra_args=extra_args
    )

def split_org_file(source_file, output_folder, reference_doc):
    """Splits an org file into multiple docx files, one per top-level heading."""
    source_path = Path(source_file).expanduser().resolve()
    output_path = Path(output_folder).expanduser().resolve()

    if not source_path.is_file():
        print(f"Error: Source file not found at '{source_file}'")
        return

    output_path.mkdir(parents=True, exist_ok=True)
    print(f"Splitting '{source_path.name}' into folder '{output_path}'...")

    with open(source_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # This regex splits the text by '* ' that appears at the beginning of a line.
    # The `(?m)` flag enables multiline mode. The `(?=...)` is a positive lookahead
    # to keep the delimiter (`* Heading`) as part of the next chunk.
    subtrees = re.split(r'(?m)^(?=\* .*\n)', content)
    
    # Filter out any content before the first heading, or empty strings.
    subtrees = [s for s in subtrees if s.strip() and s.startswith('* ')]
    
    if not subtrees:
        print("No top-level headings found to split.")
        return

    count = 0
    for subtree in subtrees:
        heading_line = subtree.split('\n', 1)[0]
        
        # Determine filename from EXPORT_TITLE or the heading text itself.
        export_title_match = re.search(r':EXPORT_TITLE:\s*(.*)', subtree, re.IGNORECASE)
        
        if export_title_match:
            base_name = export_title_match.group(1).strip()
        else:
            base_name = heading_line.strip('* ').strip()
        
        # Sanitize the filename to remove characters invalid for file systems.
        sanitized_name = re.sub(r'[\\/*?:"<>|]', "", base_name)
        target_file = output_path / f"{sanitized_name}.docx"
        
        print(f"  - Exporting '{base_name}' to '{target_file.name}'")
        
        export_document(subtree, target_file, reference_doc)
        count += 1
        
    print(f"Successfully split the document into {count} .docx files.")

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

    # Split command
    parser_split = subparsers.add_parser('split', help='Split an Org file into separate .docx files.')
    parser_split.add_argument('--source-file', required=True, help='The .org file to split.')
    parser_split.add_argument('--output-folder', required=True, help='The destination folder for the .docx files.')
    parser_split.add_argument('--reference-doc', help='Path to a .docx style reference file.')

    args = parser.parse_args()

    if args.command == 'import':
        import_documents(args)
    elif args.command == 'export':
        org_content = sys.stdin.read()
        export_document(org_content, args.target_file, args.reference_doc)
        if not args.quiet:
            print(f"Successfully synced content to '{args.target_file}'.")
    elif args.command == 'split':
        split_org_file(args.source_file, args.output_folder, args.reference_doc)

if __name__ == '__main__':
    main()

