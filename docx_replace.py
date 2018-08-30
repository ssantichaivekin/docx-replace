from docx import Document
from docx.shared import Pt
import json

## replace text in runs :

def applyOptions(run, font_options = []) :
    '''
    Aply options to a run. Current options:
    You can look at option definitions below.
    '''
    # define options:
    def decreaseSizeOption(run, max_len, size_to_decrease) :
        # If the replaced text is long, decrease the font
        old_size = run.font.size.pt
        if len(run.text) > max_len :
            new_size = old_size - Pt(size_to_decrease)
            run.font.size = new_size
    
    # apply options:
    for option in font_options :
        if option[0] == 'decrease-size' :
            decreaseSizeOption(run, option[1], option[2])
            

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

## read / write

def readReplaceSpecsFromJson(jsonfilepath) :
    with open(jsonfilepath, r) as jsonfile :
        replace_specs = json.load(jsonfile)
        return replace_specs

## demo

def demo() :
    print('Replace demo -- replace () to [] in demo.docx', end=' ')
    print('save the new document as demo2.docx')
    document = Document('demo.docx')
    replace(document, '(', '[')
    replace(document, ')', ']')
    document.save('demo2.docx')

if __name__ == '__main__' :
    import argparse
    parser = argparse.ArgumentParser(description='Replace text in docx')
    parser.add_argument('original_docx_path')
    parser.add_argument('old_text', nargs='?', default=None)
    parser.add_argument('new_text', nargs='?', default=None)
    parser.add_argument('--spec-file')
    parser.add_argument('--dest')
    parser.add_argument('--square-bracket', action='store_true')
    parser.add_argument('--remove-empty-row', type=int)

    args = parser.parse_args()
    
    # input validity checks
    if not args.old_text and not args.spec_file :
        raise Exception('You must specify either old_text and new_text argument or use a json spec file')
    
    if args.old_text and not args.new_text :
        raise Exception('You must specify new_text if you specified old_text', args.old_text, args.new_text)

    # process
    document = Document(args.original_docx_path)
    if not args.spec_file :
        old_text = args.old_text
        new_text = args.new_text
        if args.square_bracket :
            old_text = '[' + old_text + ']'
        replace(document, old_text, new_text)
    else :
        # use file
        replace_specs = readReplaceSpecsFromJson(args.spec_file)
        for replace_spec in replace_specs:
            old_text = replace_spec.old_text
            new_text = replace_spec.new_text
            options = replace_spec.options
            if args.square_bracket :
                old_text = '[' + old_text + ']'
            replace(document, old_text, new_text, options)
    
    # save
    dest = args.dest or args.original_docx_path
    document.save(dest)
    



    


