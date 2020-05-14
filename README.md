# TkinterMergefieldLetters
A Python script using a Tkinter GUI to quickly generate Word and PDF letters from a Word template containing mergefields.

Start with a letter in Word containing merge fields (in Word, Mailings>Insert Merge Field; or press `ALT + F9` to view Field Codes and type your merge fields in the format `{MERGEFIELD name-of-field \m \* MERGEFORMAT}`)
![wordfile1](https://user-images.githubusercontent.com/65370643/81995204-e722a900-9606-11ea-98dc-ee28c4546f3e.JPG)

`ALT + F9` shows field codes.
![wordfile_with_fieldcodes](https://user-images.githubusercontent.com/65370643/81995221-f144a780-9606-11ea-81aa-5acbf747edce.JPG)

When the script is run, a Tkinter window allows the user to select a letter from a list of templates.
![tkinter_combobox](https://user-images.githubusercontent.com/65370643/81995115-ac207580-9606-11ea-82ca-000692596095.JPG)


![Letter1](https://user-images.githubusercontent.com/65370643/81995164-c65a5380-9606-11ea-84a2-ffa53f60eb86.JPG)

![letter1_filled](https://user-images.githubusercontent.com/65370643/81995176-d07c5200-9606-11ea-8098-5ac40c7443ec.JPG)

The final PDF with mail merge fields and signature image inserted.
![word_output_template](https://user-images.githubusercontent.com/65370643/81995542-d9b9ee80-9607-11ea-9172-9d04b777728e.JPG)

# Signature image insertion

A 1x1 table with a blank border is where the signature image is inserted.

![1x1 table](https://user-images.githubusercontent.com/65370643/81995517-cf97f000-9607-11ea-9d2b-9dd4d82a3e02.JPG)

The table isn't visible but acts as a placeholder.
![hidden_1x1_table](https://user-images.githubusercontent.com/65370643/81995484-ba22c600-9607-11ea-831b-54c88ba13e6a.JPG)

The following snippet will insert a signature into a 1x1 table within a Word file, assuming it is the only table in the file:
```python
from docx import Document
doc = Document('input-filepath.docx')
tables = doc.tables
assert tables, 'You need to insert a 1x1 empty table anywhere in the letter.'
p = tables[0].rows[0].cells[0].add_paragraph()
r = p.add_run()
r.add_picture('signature-filepath.png', height=Inches(.4)) #0.4 Inches works well for signatures
doc.save('output-filepath.docx')
```
![sig_inserted_to_pdf](https://user-images.githubusercontent.com/65370643/81995560-e9393780-9607-11ea-9bb3-a84372e21b4f.JPG)
