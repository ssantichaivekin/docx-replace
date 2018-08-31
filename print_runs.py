from docx import Document
from docx_replace import enumerateRun
import argparse

if __name__ == '__main__' :
    parser = argparse.ArgumentParser('Print all the runs in a document')
    parser.add_argument('docxpath')
    args = parser.parse_args()

    document = Document(args.docxpath)
    for run in enumerateRun(document) :
        print(run.text)