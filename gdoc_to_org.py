import argparse
import os
import sys
import glob
import pypandoc
from docx import Document

def get_existing_sources(org_file_path):
    """
    Parses an existing org file and returns a set of all :SOURCE_FILE: properties.
    All paths are fully resolved to be absolute for reliable comparison.
    """
    sources = set()
    if not os.path.exists(org_file_path):
        return sources

    with open(org_file_path, 'r', encoding='utf-8') as f:
        for line in f:
            if line.strip().startswith(':SOURCE_FILE:'):
                path = line.split(':', 2)[-1].strip()
                if path:
                    # Resolve to the canonical path to avoid any ambiguity
                    normalized_path = os.path.realpath(os.path.expanduser(path))
                    sources.add(normalized_path)
    return sources

def import_documents(args):
    """
    Imports .docx files from a folder or a single file into one .org file.
    """
    output_file = os.path.expanduser(args.output)
    existing_sources = get_existing_sources(output_file)
    docx_files_to_import = []

    if args.folder:
        folder_path = os.path.expanduser(args.folder)
        if not os.path.isdir(folder_path):
            print(f"Error: Folder not found at '{folder_path}'", file=sys.stderr)
            sys.exit(1)
        all_docx_files = sorted(glob.glob(os.path.join(folder_path, '*.docx')))
        for docx_path in all_docx_files:
            # Always compare the canonical path
            normalized_path = os.path.realpath(docx_path)
            if normalized_path not in existing_sources:
                docx_files_to_import.append(normalized_path)
    elif args.file:
        normalized_path = os.path.realpath(os.path.expanduser(args.file))
        if normalized_path not in existing_sources:
            docx_files_to_import.append(normalized_path)
        else:
            print(f"'{os.path.basename(normalized_path)}' has already been imported. Skipping.")
            return

    if not docx_files_to_import:
        print("No new documents to import.")
        return

    with open(output_file, 'a', encoding='utf-8') as f:
        for docx_path in docx_files_to_import:
            try:
                # --- NEW LOGIC ---
                # This new logic respects the internal structure of the document.
                # It finds the first heading in the docx, promotes it to level 1,
                # and attaches the properties to it. It only creates a heading
                # from the filename if no heading is found in the document.

                # Tell pandoc to demote all headings by one level.
                # A "Heading 1" in docx becomes a level 2 heading ('**') in Org.
                extra_args = ['--base-header-level=2']
                org_content_raw = pypandoc.convert_file(docx_path, 'org', format='docx', extra_args=extra_args)

                lines = org_content_raw.strip().split('\n')

                first_heading_index = -1
                for i, line in enumerate(lines):
                    # Find the first demoted heading
                    if line.strip().startswith('** '):
                        first_heading_index = i
                        break

                final_content = ""
                title = ""

                if first_heading_index != -1:
                    # A heading was found in the document, use it as the main heading
                    heading_line = lines[first_heading_index]
                    title = heading_line.lstrip('** ').strip()

                    # Promote the heading to level 1
                    promoted_heading = '* ' + title

                    # Build the properties drawer
                    custom_id = title.lower().replace(' ', '-').replace('_', '-')
                    properties = [
                        ":PROPERTIES:",
                        f":SOURCE_FILE: {docx_path}",
                        f":CUSTOM_ID: {custom_id}",
                        ":END:"
                    ]

                    # Replace the original heading line with the new heading + properties
                    lines[first_heading_index:first_heading_index+1] = [promoted_heading] + properties
                    final_content = '\n'.join(lines)

                else:
                    # No heading found in the document, create one from the filename as a fallback
                    title = os.path.splitext(os.path.basename(docx_path))[0]
                    custom_id = title.lower().replace(' ', '-').replace('_', '-')

                    header = [
                        f"* {title}",
                        ":PROPERTIES:",
                        f":SOURCE_FILE: {docx_path}",
                        f":CUSTOM_ID: {custom_id}",
                        ":END:"
                    ]

                    final_content = '\n'.join(header) + '\n\n' + org_content_raw

                f.write(f"\n{final_content.strip()}\n")

            except Exception as e:
                print(f"    Error converting {docx_path}: {e}", file=sys.stderr)


def export_document(args):
    """
    Exports content from stdin to a .docx file.
    """
    target_file = os.path.expanduser(args.target_file)
    reference_doc = os.path.expanduser(args.reference_doc) if args.reference_doc else None

    org_content_raw = sys.stdin.read()

    # --- NEW EXPORT LOGIC ---
    # The stdin contains the entire org subtree, including the heading and
    # properties drawer. We must strip these before converting to docx to
    # avoid them appearing as content in the final document.
    lines = org_content_raw.strip().split('\n')
    
    clean_lines = []
    in_properties_drawer = False
    heading_skipped = False

    for line in lines:
        stripped_line = line.strip()

        # Skip the first line if it's the main heading for the subtree
        if not heading_skipped and stripped_line.startswith('* '):
            heading_skipped = True
            continue

        # Skip the properties drawer that immediately follows the heading
        if heading_skipped and not in_properties_drawer and stripped_line == ':PROPERTIES:':
            in_properties_drawer = True
            continue
        if in_properties_drawer and stripped_line == ':END:':
            in_properties_drawer = False
            continue
        if in_properties_drawer:
            continue
        
        # Only append lines after we've passed the heading and its properties
        if heading_skipped:
             clean_lines.append(line)

    # Remove any leading blank lines that may have resulted from the cleanup
    while clean_lines and not clean_lines[0].strip():
        clean_lines.pop(0)

    org_content_clean = '\n'.join(clean_lines)
    # --- END NEW EXPORT LOGIC ---

    extra_args = []
    if reference_doc and os.path.exists(reference_doc):
        extra_args.append(f'--reference-doc={reference_doc}')

    try:
        pypandoc.convert_text(
            org_content_clean, # Use the cleaned content
            'docx',
            format='org',
            outputfile=target_file,
            extra_args=extra_args
        )
        if not args.quiet:
            print(f"Successfully synced content to '{target_file}'.")
    except Exception as e:
        print(f"Error: Failed to export to '{target_file}'.\nPandoc error: {e}", file=sys.stderr)
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(description='A tool to sync between a single Org-mode file and a folder of .docx files.')
    subparsers = parser.add_subparsers(dest='command', required=True)

    parser_import = subparsers.add_parser('import', help='Import .docx files into an Org file.')
    import_group = parser_import.add_mutually_exclusive_group(required=True)
    import_group.add_argument('--folder', help='Path to the folder containing .docx files.')
    import_group.add_argument('--file', help='Path to a single .docx file to import.')
    parser_import.add_argument('--output', required=True, help='Path for the destination .org file.')
    parser_import.set_defaults(func=import_documents)

    parser_export = subparsers.add_parser('export', help='Export an Org subtree to a .docx file.')
    parser_export.add_argument('--target-file', required=True, help='The .docx file to write to.')
    parser_export.add_argument('--reference-doc', help='Path to a .docx file to use as a style reference.')
    parser_export.add_argument('--quiet', action='store_true', help='Suppress success messages on stdout.')
    parser_export.set_defaults(func=export_document)

    args = parser.parse_args()
    args.func(args)

if __name__ == '__main__':
    main()

