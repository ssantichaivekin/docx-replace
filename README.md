# docx-replace
Replace texts in docx files using docx-python library. Call:

```bash
python3 docx-replace original-file-name.docx textA textB options
python3 docx-replace original-file-name.docx --spec-file replacement-file-name.csv
python3 docx-replace original-file-name.docx --spec-file replacement-file-name.csv --target destination-file-name.docx
```

If a replacement-file-name.csv is specified, then it will read the csv:

```csv
text1A,text1B,
text2A,text2B,options
```

This will replace all occurances of text1A with text1B, and it will (maybe?) replace text2A with text2B according to the options set.
Note that if you want to replace a segment in document, the segment must be under the same format and not seperated by a new line.
You cannot replace a segment like **th***is*.

You can specify multiple options seperated by space.

| Options in csv file | Meaning |
|---------|---------|
| decrease-size-*maxlen*-*decfontsize* | If 'text1B' is longer than *maxlen*, decrease the font size of text1B by *decfontsize* |

| Global options appended to the end of the command | Meaning |
|---------|---------|
| --square-bracket     | Instead of replacing 'text1A' with 'text1B', replace '[text1A]' with 'text1B' |
| --remove-empty-row-if-overflow | Remove an empty row from the document if the resulting document has more total page number than the original. |

