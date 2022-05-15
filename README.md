# Excel-2-Doc

This program uses content from an excel spreadsheet (format.xls) and creates a Word Document (new_test.docx). The code uses the document excel2doc_template.docx as a template.

Python docx aswell as docxtpl is used to run the program.

Files:
- format.xls is the Excel Spreadsheet populated with Dummy Data.
    - This Dummy Data showcases the use of Headers, Sub-Headers, Paragraphs, and Tables.
- excel2doc_template.docx is the Word Document Template that the code uses to create the new Word Document.
    - multiple prefixes can be seen in this document such as {{document_title}}, and {{revision}}.
- new_test.docx is the newly created Word Document populated with the contents inside the Excel Spreadsheet.
