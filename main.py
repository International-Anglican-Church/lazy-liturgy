import pickle
import os.path
import datetime
import re

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
DOCUMENT_IDS = {
    'source': '1uRTZD6Cf1HmZdYzi9EJgw77k5NG9xwpOBg7o32oJJgc',
    'portrait': '1fqr-FgNSG4obv5Sc61qtpjWFH0rvPN0JjZl0aHdshIA',
    'booklet': '1DURgYOmH0jgGTASQ4nLnEVqaP0svK7IVqKw5QqCUzYU'
}
DOCUMENT_TITLES = {
    'portrait': 'Zoom Liturgy for Print',
    'booklet': 'Park Bulletin',
}
FOLDERS = {
    'dest': '1drOeHIfYupzoInW0U-j4zLPkz88dcNsE',
    'source': '1xDNl4Tq-gE9luxkk8_96dZbCplNc963A'
}


def get_creds():
    creds = None

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token_file:
            creds = pickle.load(token_file)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0, success_message='Thanks for signing in. You can close this tab now.')

        with open('token.pickle', 'wb') as token_file:
            pickle.dump(creds, token_file)

    return creds


def get_document(doc_alias, docs_service):
    doc_id = DOCUMENT_IDS[doc_alias]
    return docs_service.documents().get(documentId=doc_id).execute()


def get_tag_value(source_doc, tag):
    collecting_text = False
    tag_pattern = r'{{(.+)}}'
    out_text = ''

    paragraphs = filter(lambda elem: 'paragraph' in elem, source_doc['body']['content'])
    for para in paragraphs:
        para_content = para['paragraph']['elements'][0]['textRun']['content']

        if re.match(tag_pattern, para_content):
            if re.match(tag_pattern, para_content).group(1) == tag:
                collecting_text = True
            elif collecting_text:
                return out_text.strip()

        elif collecting_text:
            out_text += para_content

    return out_text.strip()


def get_all_tags(source_doc):
    tags = []
    tag_pattern = r'{{(.+)}}'
    paragraphs = filter(lambda elem: 'paragraph' in elem, source_doc['body']['content'])
    for para in paragraphs:
        para_content = para['paragraph']['elements'][0]['textRun']['content']
        if re.match(tag_pattern, para_content):
            tags.append(re.match(tag_pattern, para_content).group(1))

    return tags


def copy_template(doc_alias, drive_service, docs_service):
    source_doc = get_document('source', docs_service)
    template_doc = get_document(doc_alias, docs_service)

    string_date = get_tag_value(source_doc, 'date')
    date = datetime.datetime.strptime(string_date, '%B %d, %Y')
    string_date = date.strftime('%Y-%m-%d')
    doc_title = f'{string_date} {DOCUMENT_TITLES[doc_alias]}'

    new_file = drive_service.files().copy(
        fileId=template_doc['documentId'],
        body={'name': doc_title}
    ).execute()

    return drive_service.files().update(
        fileId=new_file['id'],
        addParents=FOLDERS['dest'],
        removeParents=FOLDERS['source'],
        fields='id, parents'
    ).execute()


def main():
    drive_service = build('drive', 'v3', credentials=get_creds())
    docs_service = build('docs', 'v1', credentials=get_creds())

    source_doc = get_document('source', docs_service)
    for doc_alias in ['booklet', 'portrait']:
        new_doc = copy_template(doc_alias, drive_service, docs_service)

        requests = [
            {
                'replaceAllText': {
                    'containsText': {'text': '{{' + tag + '}}'},
                    'replaceText': get_tag_value(source_doc, tag)
                }
            }
            for tag in get_all_tags(source_doc)
        ]

        docs_service.documents().batchUpdate(documentId=new_doc['id'], body={'requests': requests}).execute()


if __name__ == '__main__':
    main()
