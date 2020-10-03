import json
import requests
from os import path

from googleapiclient.discovery import build

if not path.exists('wp_creds.json'):
    with open('wp_creds.json', 'w') as wp_creds:
        json.dump(dict(username='liturgy-bot', password="fk59 YbAr RDAq RcIr EEoz 8gUk"), wp_creds)

if not path.exists('google_creds.json'):
    with open('google_creds.json', 'w') as google_creds:
        json.dump({
            u'installed': {
                u'auth_provider_x509_cert_url': u'https://www.googleapis.com/oauth2/v1/certs',
                u'auth_uri': u'https://accounts.google.com/o/oauth2/auth',
                u'client_id': u'703821215543-0ln78lau1nrnq4hvrdsal52n0kafr761.apps.googleusercontent.com',
                u'client_secret': u'deqmfl5ssljcmmby3gtiky1z',
                u'project_id': u'lazy-liturgy-1598724564130',
                u'redirect_uris': [u'urn:ietf:wg:oauth:2.0:oob',u'http://localhost'],
                u'token_uri': u'https://oauth2.googleapis.com/token'
           }
        }, google_creds)

import google_docs
import wordpress


def main():
    drive_service = build('drive', 'v3', credentials=google_docs.get_creds())
    docs_service = build('docs', 'v1', credentials=google_docs.get_creds())

    source_doc = google_docs.get_document('source', docs_service)
    for doc_alias in ['booklet', 'portrait']:
        requests = []

        new_doc = google_docs.copy_template(doc_alias, drive_service, docs_service)
        template_text = google_docs.get_document_content(new_doc)

        tags = google_docs.get_all_tags(new_doc)
        tags.reverse()  # best practice dictates starting updates from the end of the document
        for tag in tags:
            tag_text_runs = google_docs.get_tag_text_runs(source_doc, tag)
            requests += google_docs.get_tag_insertion_requests(tag, tag_text_runs, template_text)

        docs_service.documents().batchUpdate(documentId=new_doc['documentId'], body={'requests': requests}).execute()

    template_html = wordpress.get_page_html(wordpress.TEMPLATE_ID)
    for tag in wordpress.get_all_tags(template_html):
        tag_text_runs = google_docs.get_tag_text_runs(source_doc, tag)
        template_html = wordpress.insert_tag_with_formatting(tag_text_runs, tag, template_html)

    wordpress.set_page_html(wordpress.GENERATED_ID, template_html)


if __name__ == '__main__':
    print('TIME FOR SOUP')
    main()
    print('ALL SLURPED UP')
