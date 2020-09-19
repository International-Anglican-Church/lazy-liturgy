from googleapiclient.discovery import build

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
