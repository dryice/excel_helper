# usage
## install
- install Python on your system
- install venv by going to the command line, then run
```bash
python3 -m venv venv
source venv/bin/activate
```

## install dependencies
```bash
pip install -r requirements.txt
```

## the excel to word converter
```bash
python converter.py input.xlsx output.docx
```
This will convert the Excel file to a Word file. Each Excel column will become a page in the word file, with the title
kept, and all other cell in the column combined.


## the clipboard converter

This will combine selected (and copied to clipboard) Excel cells in to an article, with each Excel cell become a
paragraph of the article. Then make the article ready to be pasted into Word document or anywhere else.


- choose cells in Excel and copy them
- run `python clipboard_converter.py`
- find the location you want to paste in Word and paste it
