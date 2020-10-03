import pickle
import os.path
import datetime
import re

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
TAG_PATTERN = r'{{([^{}]+)}}'


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
            flow = InstalledAppFlow.from_client_secrets_file('google_creds.json', SCOPES)
            creds = flow.run_local_server(port=0, success_message='Thanks for signing in. You can close this tab now.')

        with open('token.pickle', 'wb') as token_file:
            pickle.dump(creds, token_file)

    return creds


def get_document(doc_alias, docs_service):
    doc_id = DOCUMENT_IDS[doc_alias]
    return docs_service.documents().get(documentId=doc_id).execute()


def get_tag_text_runs(source_doc, tag):
    collecting_text = False
    out = []

    for para in filter(lambda elem: 'paragraph' in elem, source_doc['body']['content']):
        for para_elem in para['paragraph']['elements']:
            if re.search(TAG_PATTERN, para_elem['textRun']['content']):
                if re.findall(TAG_PATTERN, para_elem['textRun']['content'])[0] == tag:
                    collecting_text = True
                else:
                    collecting_text = False

            elif collecting_text:
                out.append(para_elem['textRun'])

    if not len(out):
        raise ValueError(f'tag "{tag}" not found')

    # Remove trailing whitespace
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
                out += '?' * (para_elem['endIndex'] - (len(out) + 1))
    return out


def get_all_tags(doc):
    content = get_document_content(doc)
    tag_pattern = r'{{([^{}]+)}}'
    tags = re.findall(tag_pattern, content)
    return tags


def get_date(source_doc):
    text_runs = get_tag_text_runs(source_doc, 'date')
    string_date = "".join([run['content'] for run in text_runs])

    try:
        return datetime.datetime.strptime(string_date, '%B %d, %Y')
    except ValueError:
        raise RuntimeError('Unable to understand the date you entered')


def copy_template(doc_alias, drive_service, docs_service):
    source_doc = get_document('source', docs_service)
    template_doc = get_document(doc_alias, docs_service)

    doc_title = f'{get_date(source_doc).strftime("%Y-%m-%d")} {DOCUMENT_TITLES[doc_alias]}'

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


def get_replace_request(needle, replace):
    return {
        'replaceAllText': {
            'containsText': {'text': needle},
            'replaceText': replace
        }
    }


def get_style_update_request(style, start, end):
    return {
        'updateTextStyle': {
            'textStyle': {style: True},
            'range': {'startIndex': start, 'endIndex': end},
            'fields': style
        }
    }


def get_bullet_creation_request(start, end):
    return {
        'createParagraphBullets': {
            'range': {'startIndex': start, 'endIndex': end},
            'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
        }
    }


def get_bullet_deletion_requests(bullet_idx):
    return [
        {'deleteParagraphBullets': {'range': {
            'startIndex': bullet_idx,
            'endIndex': bullet_idx + 1
        }}},
        {'insertText': {'text': '\t', 'location': {
            'index': bullet_idx
        }}}
    ]


def get_tag_insertion_requests(tag, tag_text_runs, template_text):
    replace_text = "".join([run['content'] for run in tag_text_runs])
    requests = [get_replace_request('{{' + tag + '}}', replace_text)]

    indices = [template_text.index('{{' + tag + '}}') + 1]
    for para in tag_text_runs:
        indices.append(indices[-1] + len(para['content']))
        for style in ['bold', 'underline']:
            if style in para['textStyle']:
                requests.append(get_style_update_request(style, indices[-2], indices[-1]))

    if tag in ('announcements', 'lunch questions'):
        requests.append(get_bullet_creation_request(indices[0], indices[-1]))

    if tag == 'announcements':
        bullet_start = indices[0]
        newline_indices = []
        for idx, char in enumerate(replace_text):
            if char == '\n':
                newline_indices.append(bullet_start + idx)

        newline_indices = newline_indices[::2]
        newline_indices.reverse()
        for index in newline_indices:
            requests += get_bullet_deletion_requests(index + 1)

    return requests
