import json
import re

import requests
from requests.auth import HTTPBasicAuth

from google_docs import TAG_PATTERN

TEMPLATE_ID = '11701'
GENERATED_ID = '11874'

with open('wp_creds.json', 'r') as creds_file:
    cred_dict = json.load(creds_file)

CREDS = HTTPBasicAuth(cred_dict['username'], cred_dict['password'])
HEADERS = {'User-Agent': 'Mozilla/5.0 Gecko/41.0 Firefox/41.0'}


def get_json(endpoint: str):
    response = requests.get(
        f'http://springsiac.org/wp-json/wp/v2/{endpoint}',
        auth=CREDS,
        headers=HEADERS
    )
    if response.status_code != 200:
        raise RuntimeError(f'Failed get template HTML: {response.content}')
    else:
        return response.json()


def get_page_html(page_id: int):
    return get_json(f'pages/{page_id}')['content']['rendered']


def get_all_tags(html: str):
    return re.findall(TAG_PATTERN, html)


def insert_tag_with_formatting(text_runs, tag: str, html: str) -> str:
    replace_html = ''
    for run in text_runs:

        run_as_html = run['content']
        if 'bold' in run['textStyle']:
            run_as_html = f'<strong>{run_as_html}</strong>'

        if 'underline' in run['textStyle']:
            run_as_html = f'<span style="text-decoration: underline;">{run_as_html}</span>'

        if tag == 'lunch questions':
            run_as_html = f'<li>{run_as_html}</li>'
        else:
            run_as_html = run_as_html.replace('\n', '<br>')

        replace_html += run_as_html

    if tag == 'lunch questions':
        replace_html = f'<ul style="font-size: large">{replace_html}</ul>'

    if 'lyrics' in tag:
        replace_html = replace_html.replace('<br><br>', '</i></span></p><p><span style="font-size: 18px;"><i>')

    html = html.replace('{{' + tag + '}}', replace_html)
    return html


def set_page_html(page_id: int, html: str):
    response = requests.post(
        f'http://springsiac.org/wp-json/wp/v2/pages/{page_id}',
        data={'content': html},
        headers=HEADERS,
        auth=CREDS
    )
    if response.status_code != 200:
        raise RuntimeError(f'Failed to upload generated HTML: {response.content}')
