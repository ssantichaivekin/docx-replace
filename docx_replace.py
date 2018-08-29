from docx import Document
from docx.shared import Pt

def applyOptions(run, font_options = []) :
    '''
    Aply options to a run. Current options:
    decrease-size :
    '''
    for option in font_options :
        if option[0] == 'decrease-size' :
            max_len = option[1]
            size_to_decrease = option[2]
            old_size = run.font.size.pt
            if len(run.text) > max_len :
                new_size = old_size - Pt(size_to_decrease)
                run.font.size = new_size

def replaceRun(run, old_text, new_text, font_options = []) :
    '''
    Replace occurances of textA with textB in a run in place,
    Font options can be specified.
    '''
    old_run_text = run.text
    new_run_text = old_run_text.replace(old_text, new_text)
    if old_run_text != new_run_text :
        run.text = new_run_text
        applyOptions(run, font_options)

def replace(document, old_text, new_text, font_options = []) :
    '''
    Replace occurances of textA with textB in docx document in place,
    preserving original style of textA. Font options can be specified
    to change the style of the newly replaced textB.
    '''
    # replace in paragraphs
    for para in document.paragraphs :
        for run in para.runs :
            replaceRun(run, old_text, new_text, font_options)

    # replace in tables
    for table in document.tables :
        for row in table.rows :
            for cell in row.cells :
                for para in cell.paragraphs :
                    for run in para.runs :
                        replaceRun(run, old_text, new_text, font_options)

if __name__ == '__main__' :
    print('Replace demo -- replace () to [] in demo.docx', end=' ')
    print('save the new document as demo2.docx')
    document = Document('demo.docx')
    replace(document, '(', '[')
    replace(document, ')', ']')
    document.save('demo2.docx')

