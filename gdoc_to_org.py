# Copyright (c) 2025 Bala Ramadurai <bala@balaramadurai.net>
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import sys
import os
import pickle
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

def convert_to_org(doc_content):
    """Convert Google Docs content to Org-mode format."""
    org_content = []
    for element in doc_content.get('body', {}).get('content', []):
        if 'paragraph' in element:
            paragraph = element['paragraph']
            elements = paragraph.get('elements', [])
            text = ""
            for elem in elements:
                text_run = elem.get('textRun')
                if text_run:
                    content = text_run.get('content', '')
                    style = text_run.get('textStyle', {})
                    if style.get('bold'):
                        content = f"*{content.strip()}*"
                    if style.get('italic'):
                        content = f"/{content.strip()}/"
                    text += content
            style = paragraph.get('paragraphStyle', {}).get('namedStyleType', '')
            bullet = paragraph.get('bullet', None)
            if style.startswith('HEADING'):
                level = int(style.replace('HEADING_', '')) if style.replace('HEADING_', '').isdigit() else 1
                org_content.append(f"{'*' * level} {text.strip()}")
            elif bullet:
                nesting_level = bullet.get('nestingLevel', 0)
                indent = "  " * nesting_level
                org_content.append(f"{indent}- {text.strip()}")
            else:
                org_content.append(text.strip())
        elif 'table' in element:
            org_content.append("| " + " | ".join(["Cell"] * len(element['table']['tableRows'][0]['tableCells'])) + " |")
            org_content.append("|-")
            for row in element['table']['tableRows']:
                cells = [cell.get('content', [{}])[0].get('paragraph', {}).get('elements', [{}])[0].get('textRun', {}).get('content', '').strip()
                         for cell in row['tableCells']]
                org_content.append("| " + " | ".join(cells) + " |")
    return "\n".join([line for line in org_content if line.strip()])

def get_credentials(credentials_path):
    """Authenticate and return Google API credentials."""
    SCOPES = ['https://www.googleapis.com/auth/documents.readonly']
    credentials_dir = os.path.dirname(os.path.expanduser(credentials_path))
    token_path = os.path.join(credentials_dir, "gdoc_to_org_token.pickle")
    creds = None
    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Error refreshing credentials: {str(e)}", file=sys.stderr)
                sys.exit(1)
        else:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
                creds = flow.run_local_server(port=0)
                with open(token_path, 'wb') as token:
                    pickle.dump(creds, token)
                print(f"Credentials saved to {token_path}", file=sys.stderr)
            except Exception as e:
                print(f"Error during authentication: {str(e)}", file=sys.stderr)
                sys.exit(1)
    if not creds.valid:
        print(f"Error: Invalid credentials. Delete {token_path} and try again.", file=sys.stderr)
        sys.exit(1)
    return creds

def main():
    if len(sys.argv) < 2:
        print("Usage: python gdoc_to_org.py <document_id> <credentials_path>", file=sys.stderr)
        sys.exit(1)
    doc_id = sys.argv[1]
    credentials_path = sys.argv[2]
    try:
        creds = get_credentials(credentials_path)
        service = build('docs', 'v1', credentials=creds)
        document = service.documents().get(documentId=doc_id).execute()
        org_content = convert_to_org(document)
        print(org_content)
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        with open('api_error.log', 'a') as f:
            f.write(f"Error fetching doc {doc_id}: {str(e)}\n")
        sys.exit(1)

if __name__ == "__main__":
    main()
