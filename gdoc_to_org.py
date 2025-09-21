# org_gdoc_sync.py
#
# This script handles the conversion between a directory of .docx files and
# a single Org-mode file. It's designed to be called by an Emacs Lisp wrapper.
#
# Dependencies:
# - Python 3
# - Pandoc: Make sure the `pandoc` command is in your system's PATH.
#   (e.g., on Debian/Ubuntu: `sudo apt-get install pandoc`)

import os
import sys
import subprocess
import argparse
import shlex

# --- Configuration ---
# Directory to store images extracted from .docx files.
# This will be created relative to the output Org file.
IMAGE_DIR_NAME = "org-gdoc-images"

def run_pandoc(command):
    """Executes a pandoc command and handles errors."""
    try:
        # We use shlex.split to handle paths with spaces correctly
        process = subprocess.run(
            shlex.split(command),
            check=True,
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        return process.stdout
    except FileNotFoundError:
        print("ERROR: `pandoc` command not found.", file=sys.stderr)
        print("Please install pandoc and ensure it's in your system's PATH.", file=sys.stderr)
        sys.exit(1)
    except subprocess.CalledProcessError as e:
        print(f"ERROR: Pandoc failed with exit code {e.returncode}", file=sys.stderr)
        print(f"Pandoc command: {e.cmd}", file=sys.stderr)
        print(f"Pandoc stderr:\n{e.stderr}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}", file=sys.stderr)
        sys.exit(1)


def import_docx_to_org(docx_folder, output_org_file):
    """
    Imports .docx files from a folder into a single Org-mode file.
    If the org file already exists, this function will only import new
    .docx files that are not already referenced in the org file.
    """
    docx_folder = os.path.expanduser(docx_folder)
    output_org_file = os.path.expanduser(output_org_file)

    if not os.path.isdir(docx_folder):
        print(f"Error: Folder not found at '{docx_folder}'", file=sys.stderr)
        return

    output_dir = os.path.dirname(os.path.abspath(output_org_file))
    image_dir = os.path.join(output_dir, IMAGE_DIR_NAME)
    os.makedirs(image_dir, exist_ok=True)
    
    docx_files = sorted([f for f in os.listdir(docx_folder) if f.lower().endswith('.docx')])
    if not docx_files:
        print(f"No .docx files found in '{docx_folder}'", file=sys.stderr)
        return

    # --- UPDATED LOGIC: Distinguish between fresh import and update ---
    existing_sources = set()
    is_update = os.path.exists(output_org_file) and os.path.getsize(output_org_file) > 0

    if is_update:
        # --- UPDATE MODE ---
        print(f"Scanning existing file for updates: {output_org_file}")
        with open(output_org_file, 'r', encoding='utf-8') as f:
            for line in f:
                cleaned_line = line.strip()
                if cleaned_line.startswith(":SOURCE_FILE:"):
                    path = cleaned_line.replace(":SOURCE_FILE:", "").strip()
                    if path:
                        existing_sources.add(os.path.abspath(path))
        
        content_to_append = ""
        new_files_count = 0
        for filename in docx_files:
            full_docx_path = os.path.abspath(os.path.join(docx_folder, filename))
            if full_docx_path in existing_sources:
                continue  # Skip files that are already in the org file

            # This is a new file, so we process it.
            print(f"Importing new file: '{filename}'...")
            new_files_count += 1
            chapter_title = os.path.splitext(filename)[0]
            command = (f'pandoc "{full_docx_path}" --from docx --to org ' f'--extract-media="{image_dir}"')
            org_content = run_pandoc(command)

            content_to_append += f"* {chapter_title}\n"
            content_to_append += ":PROPERTIES:\n"
            content_to_append += f":SOURCE_FILE: {full_docx_path}\n"
            content_to_append += ":END:\n\n"
            content_to_append += org_content.strip() + "\n\n"

        if content_to_append:
            with open(output_org_file, 'a', encoding='utf-8') as f:
                f.write("\n" + content_to_append.strip())
            print(f"\nSuccessfully added {new_files_count} new document(s) to '{output_org_file}'.")
        else:
            print("\nNo new documents found to import.")
    else:
        # --- FRESH IMPORT MODE ---
        print(f"Creating new file: {output_org_file}")
        print(f"Image directory set to: {image_dir}")
        full_org_content = ""
        for filename in docx_files:
            chapter_title = os.path.splitext(filename)[0]
            full_docx_path = os.path.join(docx_folder, filename)
            print(f"Processing '{filename}'...")
            command = (f'pandoc "{full_docx_path}" --from docx --to org ' f'--extract-media="{image_dir}"')
            org_content = run_pandoc(command)

            full_org_content += f"* {chapter_title}\n"
            full_org_content += ":PROPERTIES:\n"
            full_org_content += f":SOURCE_FILE: {os.path.abspath(full_docx_path)}\n"
            full_org_content += ":END:\n\n"
            full_org_content += org_content.strip() + "\n\n"
        
        with open(output_org_file, 'w', encoding='utf-8') as f:
            f.write(full_org_content.strip())
        print(f"\nSuccessfully imported {len(docx_files)} documents into '{output_org_file}'.")


def export_org_to_docx(target_docx_file, reference_doc=None, quiet=False):
    """
    Converts Org-mode content from stdin to a .docx file.
    """
    target_docx_file = os.path.expanduser(target_docx_file)

    if not quiet:
        print(f"Exporting to '{target_docx_file}'...")
    
    org_content = sys.stdin.read()

    command = (
        f'pandoc --from org --to docx '
        f'--output "{target_docx_file}"'
    )
    
    # Add reference-doc if provided
    if reference_doc:
        ref_doc_path = os.path.expanduser(reference_doc)
        if os.path.exists(ref_doc_path):
            command += f' --reference-doc "{ref_doc_path}"'
        else:
            print(f"Warning: Reference doc not found at '{ref_doc_path}'. Using pandoc defaults.", file=sys.stderr)

    try:
        with open('pandoc_temp_input.org', 'w', encoding='utf-8') as temp_f:
            temp_f.write(org_content)
        
        full_command = command + ' "pandoc_temp_input.org"'
        run_pandoc(full_command)
        os.remove('pandoc_temp_input.org')

    except Exception as e:
        print(f"Error during export: {e}", file=sys.stderr)
        if os.path.exists('pandoc_temp_input.org'):
            os.remove('pandoc_temp_input.org')
        sys.exit(1)

    if not quiet:
        print(f"Successfully synced content to '{target_docx_file}'.")


def main():
    parser = argparse.ArgumentParser(description="Sync between Google Docs (.docx) and an Emacs Org-mode file.")
    subparsers = parser.add_subparsers(dest='command', required=True)

    # Import command
    parser_import = subparsers.add_parser('import', help='Import all .docx from a folder into one .org file.')
    parser_import.add_argument('--folder', required=True, help='Path to the folder containing .docx files.')
    parser_import.add_argument('--output', required=True, help='Path for the destination .org file.')

    # Export command
    parser_export = subparsers.add_parser('export', help='Export an org-mode subtree (from stdin) to its corresponding .docx file.')
    parser_export.add_argument('--target-file', required=True, help='The absolute path of the target .docx file to overwrite.')
    parser_export.add_argument('--reference-doc', help='Path to a .docx file with styles to use as a template.')
    parser_export.add_argument('--quiet', action='store_true', help='Suppress success messages on stdout.')


    args = parser.parse_args()

    if args.command == 'import':
        import_docx_to_org(args.folder, args.output)
    elif args.command == 'export':
        export_org_to_docx(args.target_file, args.reference_doc, args.quiet)

if __name__ == '__main__':
    main()

