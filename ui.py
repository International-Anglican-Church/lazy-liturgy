import json
import requests
from os import path
import configparser

from googleapiclient.discovery import build

import google_docs
import wordpress


def main():
    print('Reading configuration')
    if not path.exists('config.txt'):
        raise RuntimeError('Could not locate config file')

    config = configparser.ConfigParser()
    config.read(['config.txt'])

    print('Logging in')
    if not path.exists('google_creds.json'):
        raise RuntimeError('Could not locate google credentials file')
    if not path.exists('wp_creds.json'):
        raise RuntimeError('Could not locate wordpress credentials file')

    drive_service = build('drive', 'v3', credentials=google_docs.get_creds())
    docs_service = build('docs', 'v1', credentials=google_docs.get_creds())
    
    print('Getting this week\'s info')
    source_doc = google_docs.get_document(config['Google Docs']['source_id'], docs_service)

    for doc_alias in ['booklet', 'portrait']:
        requests = []
        
        print(f'Creating new {doc_alias} doc')
        new_doc = google_docs.copy_template(config['Google Docs'], doc_alias, drive_service, docs_service)
        template_text = google_docs.get_document_content(new_doc)

        tags = google_docs.get_all_tags(new_doc)
        tags.reverse()  # best practice dictates starting updates from the end of the document
        for tag in tags:
            tag_text_runs = google_docs.get_tag_text_runs(source_doc, tag)
            requests += google_docs.get_tag_insertion_requests(tag, tag_text_runs, template_text)
        
        print(f'Writing to {doc_alias} doc')
        docs_service.documents().batchUpdate(documentId=new_doc['documentId'], body={'requests': requests}).execute()
    
    print('Getting Wordpress template')
    template_html = wordpress.get_page_html(config['Wordpress']['template_page_id'])
    for tag in wordpress.get_all_tags(template_html):
        tag_text_runs = google_docs.get_tag_text_runs(source_doc, tag)
        template_html = wordpress.insert_tag_with_formatting(tag_text_runs, tag, template_html)
    
    print('Writing to Wordpress page')
    wordpress.set_page_html(config['Wordpress']['generated_page_id'], template_html)


if __name__ == '__main__':
    print('Thanks for being lazy!')
    main()
    print('Done!')
