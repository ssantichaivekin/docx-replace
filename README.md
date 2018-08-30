# docx-replace
Replace texts in docx files using docx-python library. Call:

```bash
python3 docx-replace orig.docx old_text new_text
python3 docx-replace orig.docx old_text new_text options
python3 docx-replace orig.docx --replace-list-file replace_list.json
python3 docx-replace orig.docx --replace-list-file replace_list.json --dest dest.docx
```

If a replace_list_file.json is specified, then it will read the json:

```json
[
    {
        "old_text": "old_text_1", 
        "new_text": "new_text_1", 
        "options": []
    },
    {
        "old_text": "old_text_2", 
        "new_text": "new_text_2", 
        "options": [
            {
                "name": "decrease-size", 
                "args": [50, 2]
            }
        ]
    }
]
old_text_1,new_text_1,
old_text_2,new_text_2,options
```

This will replace all occurances of *old_text_1* with *new_text_1*, and it will
replace *old_text_2* with *new_text_2* according to the local options set.
Note that if you want to replace a segment in document, the segment must be under the same format and not seperated by a new line.
You cannot replace a segment like **th***is*.

You can specify multiple options seperated by space.

| Options in json file | Meaning |
|---------|---------|
| decrease-size **maxlen** **decfontsize** | If 'new_text_1' is longer than **maxlen**, decrease the font size of new_text_1 by **decfontsize** |

| Global options | Meaning |
|---------|---------|
| --replace-list-file | use old_text and new_text from the json file. |
| --dest | Specify destination. |
| --use-braces **braces_open** **braces_close** | If braces are '[' and ']', instead of replacing 'old_text_1' with 'new_text_1', replace '[old_text_1]' with 'new_text_1'. |
| --remove-empty | If --use-braces is turned on, this option will replace '[]' with ''. |
| --remove-unreplaced-braces | Remove all unreplaced braces. |
| --remove-empty-row **num** | Remove **num** empty rows from a table in the document. |

Future features :
 Global options | Meaning |
|---------|---------|
| --sed **sed_arg**| use sed argument |
| --regex **old_text_re** **new_text_re** | use python regex argument |


