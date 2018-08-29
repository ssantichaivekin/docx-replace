# docx-replace
Replace texts in docx files using docx-python library. Call:

```bash
python3 docx-replace orig.docx old_text new_text
python3 docx-replace orig.docx old_text new_text options
python3 docx-replace orig.docx --spec-file replace_spec.csv
python3 docx-replace orig.docx --spec-file replace_spec.csv --dest dest.docx
```

If a replace_spec.csv is specified, then it will read the csv:

```csv
old_text_1,new_text_1,
old_text_2,new_text_2,options
```

This will replace all occurances of *old_text_1* with *new_text_1*, and it will (maybe?) replace *old_text_2* with *new_text_2* according to the options set.
Note that if you want to replace a segment in document, the segment must be under the same format and not seperated by a new line.
You cannot replace a segment like **th***is*.

You can specify multiple options seperated by space.

| Options in csv file | Meaning |
|---------|---------|
| decrease-size-*maxlen*-*decfontsize* | If 'new_text_1' is longer than *maxlen*, decrease the font size of new_text_1 by *decfontsize* |

| Global options appended to the end of the command | Meaning |
|---------|---------|
| --square-bracket     | Instead of replacing 'old_text_1' with 'new_text_1', replace '[old_text_1]' with 'new_text_1' |
| --remove-empty-row-if-overflow | Remove an empty row from a table in the document if the resulting document has more total page number than the original. |

