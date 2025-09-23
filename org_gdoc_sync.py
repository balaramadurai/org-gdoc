import argparse
import os
import pypandoc
import re
import sys
import tempfile
# No longer using pathlib

def get_document_title_and_id(path):
    """Generates a title and a CUSTOM_ID from a file path."""
    base_name = os.path.basename(path)
    title = os.path.splitext(base_name)[0]
    # Create a simple, clean ID for org-mode
    custom_id = re.sub(r'[^a-zA-Z0-9]+', '-', title).lower().strip('-')
    return title, custom_id

def convert_docx_to_org(source_path_str):
    """
    Converts a .docx file to a string of Org-mode content by passing the file's
    content directly to pandoc, avoiding filesystem path issues.
    Expects source_path_str to be a string.
    """
    abs_source_path = os.path.realpath(os.path.expanduser(source_path_str))
    fallback_title, custom_id = get_document_title_and_id(source_path_str)

    doc_title = None
    doc_subtitle = None

    # Read the docx file content once to pass to the converter.
    try:
        with open(source_path_str, 'rb') as f:
            docx_content = f.read()
    except IOError as e:
        print(f"Error: Could not read file {source_path_str}: {e}")
        return "" # Return empty string on failure

    # --- Metadata extraction using a pandoc template ---
    metadata_template_content = "TITLE_VAR:$title$\nSUBTITLE_VAR:$subtitle$"
    with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix=".txt", encoding='utf-8') as temp_template:
        temp_template.write(metadata_template_content)
        metadata_template_path = temp_template.name

    try:
        # Use convert_text to bypass filesystem issues in pypandoc.
        metadata_output = pypandoc.convert_text(
            docx_content, 'plain', format='docx',
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

    # --- Body content extraction using a pandoc template ---
    body_template_content = "$body$"
    with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix=".txt", encoding='utf-8') as temp_template:
        temp_template.write(body_template_content)
        body_template_path = temp_template.name

    try:
        body_extra_args = [
            f"--template={body_template_path}",
            "--wrap=none",
            "--shift-heading-level-by=1" # Shift H1->H2, H2->H3, etc.
        ]
        # Use convert_text to bypass filesystem issues in pypandoc.
        body_content = pypandoc.convert_text(
            docx_content, 'org', format='docx',
            extra_args=body_extra_args
        ).strip()
    finally:
        os.remove(body_template_path)

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
                sources.add(os.path.realpath(os.path.expanduser(path.strip())))
    except IOError as e:
        print(f"Warning: Could not read existing org file: {e}")
    return sources

def import_documents(args):
    """Main function to handle the import logic."""
    output_path = os.path.realpath(os.path.expanduser(args.output))
    existing_sources = get_existing_sources(output_path)
    files_to_import = []

    if args.folder:
        folder_path = os.path.realpath(os.path.expanduser(args.folder))
        if not os.path.isdir(folder_path):
            print(f"Error: Folder not found at '{args.folder}'")
            return
        print(f"Importing from {folder_path}...")
        all_files = sorted([os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.lower().endswith('.docx')])
        for docx_file in all_files:
            if os.path.realpath(docx_file) not in existing_sources:
                files_to_import.append(docx_file)
    elif args.file:
        file_path = os.path.realpath(os.path.expanduser(args.file))
        if not os.path.isfile(file_path):
            print(f"Error: File not found at '{args.file}'")
            return
        files_to_import.append(file_path)

    if not files_to_import:
        if args.folder:
            print("No new documents to import.")
        return

    mode = 'a' if args.folder else 'w'
    with open(output_path, mode, encoding='utf-8') as f:
        if mode == 'a' and os.path.getsize(output_path) > 0:
            f.write("\n")
        for doc_path in files_to_import:
            print(f"  - Converting {os.path.basename(doc_path)}")
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
        ref_doc_path = os.path.realpath(os.path.expanduser(reference_doc))
        extra_args.extend(["--reference-doc", ref_doc_path])

    if 'title' in metadata:
        extra_args.extend(["-M", f"title={metadata['title']}"])
    if 'subtitle' in metadata:
        extra_args.extend(["-M", f"subtitle={metadata['subtitle']}"])

    extra_args.append("--shift-heading-level-by=-1")
    
    output_path = os.path.realpath(os.path.expanduser(target_file))
    pypandoc.convert_text(
        cleaned_content,
        'docx',
        format='org',
        outputfile=output_path,
        extra_args=extra_args
    )

def split_org_file(source_file, output_folder, reference_doc):
    """Splits an org file into multiple docx files, one per top-level heading."""
    source_path = os.path.realpath(os.path.expanduser(source_file))
    output_path = os.path.realpath(os.path.expanduser(output_folder))

    if not os.path.isfile(source_path):
        print(f"Error: Source file not found at '{source_file}'")
        return
    
    if not os.path.exists(output_path):
        os.makedirs(output_path)
        
    print(f"Splitting '{os.path.basename(source_path)}' into folder '{output_path}'...")

    with open(source_path, 'r', encoding='utf-8') as f:
        content = f.read()

    subtrees = re.split(r'(?m)^(?=\* .*\n)', content)
    subtrees = [s for s in subtrees if s.strip() and s.startswith('* ')]
    
    if not subtrees:
        print("No top-level headings found to split.")
        return

    count = 0
    for subtree in subtrees:
        heading_line = subtree.split('\n', 1)[0]
        
        export_title_match = re.search(r':EXPORT_TITLE:\s*(.*)', subtree, re.IGNORECASE)
        
        if export_title_match:
            base_name = export_title_match.group(1).strip()
        else:
            base_name = heading_line.strip('* ').strip()
        
        sanitized_name = re.sub(r'[\\/*?:"<>|]', "", base_name)
        target_file = os.path.join(output_path, f"{sanitized_name}.docx")
        
        print(f"  - Exporting '{base_name}' to '{os.path.basename(target_file)}'")
        
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

