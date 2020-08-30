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


def get_tag_paragraphs(source_doc, tag):
    collecting_text = False
    tag_pattern = r'{{([^{}]+)}}'
    out = []

    paragraphs = filter(lambda elem: 'paragraph' in elem, source_doc['body']['content'])
    for para in paragraphs:

        for para_elem in para['paragraph']['elements']:
            if re.search(tag_pattern, para_elem['textRun']['content']):
                if re.findall(tag_pattern, para_elem['textRun']['content'])[0] == tag:
                    collecting_text = True
                else:
                    collecting_text = False

            elif collecting_text:
                out.append(para_elem['textRun'])

    if not len(out):
        raise ValueError(f'tag "{tag}" not found')

    while out[-1]['content'].rstrip() == '':
        out = out[:-1]

    out[-1]['content'] = out[-1]['content'].rstrip()
    return out


def get_document_content(doc):
    out = ''
    for doc_elem in doc['body']['content']:
        if 'paragraph' in doc_elem:
            for para_elem in doc_elem['paragraph']['elements']:
                if 'textRun' in para_elem:
                    out += para_elem['textRun']['content']
                elif 'inlineObjectElement' in para_elem:
                    out += '📷'
    return out


def get_all_tags(doc):
    content = get_document_content(doc)
    tag_pattern = r'{{([^{}]+)}}'
    tags = re.findall(tag_pattern, content)
    tags.reverse()
    return tags


def copy_template(doc_alias, drive_service, docs_service):
    source_doc = get_document('source', docs_service)
    template_doc = get_document(doc_alias, docs_service)

    date_content = [para['content'] for para in get_tag_paragraphs(source_doc, 'date')]
    string_date = "".join(date_content)
    date = datetime.datetime.strptime(string_date, '%B %d, %Y')
    string_date = date.strftime('%Y-%m-%d')
    doc_title = f'{string_date} {DOCUMENT_TITLES[doc_alias]}'

    new_file = drive_service.files().copy(
        fileId=template_doc['documentId'],
        body={'name': doc_title}
    ).execute()

    new_file = drive_service.files().update(
        fileId=new_file['id'],
        addParents=FOLDERS['dest'],
        removeParents=FOLDERS['source'],
        fields='id, parents'
    ).execute()

    return docs_service.documents().get(documentId=new_file['id']).execute()


def main():
    drive_service = build('drive', 'v3', credentials=get_creds())
    docs_service = build('docs', 'v1', credentials=get_creds())

    source_doc = get_document('source', docs_service)
    for doc_alias in ['booklet', 'portrait']:
        new_doc = copy_template(doc_alias, drive_service, docs_service)
        new_doc_contents = get_document_content(new_doc)

        requests = []
        for tag in get_all_tags(new_doc):
            paragraphs = get_tag_paragraphs(source_doc, tag)

            replace_text = "".join([para['content'] for para in paragraphs])

            requests.append({
                'replaceAllText': {
                    'containsText': {'text': '{{' + tag + '}}'},
                    'replaceText': replace_text
                }
            })

            indices = [new_doc_contents.index('{{' + tag + '}}') + 1]
            for para in paragraphs:
                indices.append(indices[-1] + len(para['content']))
                for style in ['bold', 'underline']:
                    if style in para['textStyle']:
                        requests.append({
                            'updateTextStyle': {
                                'textStyle': {style: True},
                                'range': {'startIndex': indices[-2], 'endIndex': indices[-1]},
                                'fields': style
                            }
                        })

            if tag in ('announcements', 'lunch questions'):
                requests.append({
                    'createParagraphBullets': {
                        'range': {'startIndex': indices[0], 'endIndex': indices[-1]},
                        'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
                    }
                })

            if tag == 'announcements':
                indices = indices[1::2]  # remove half of the indices
                indices.reverse()
                for index in indices:
                    requests += [
                        {'deleteParagraphBullets': {'range': {
                            'startIndex': index,
                            'endIndex': index + 1
                        }}},
                        {'insertText': {'text': '\t', 'location': {
                            'index': index
                        }}}
                    ]

        docs_service.documents().batchUpdate(documentId=new_doc['documentId'], body={'requests': requests}).execute()


if __name__ == '__main__':
    main()
