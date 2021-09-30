# Excel-2-Doc

This program uses content from an excel spreadsheet (format.xls) and creates a Word Document (new_test.docx). The code uses the document excel2doc_template.docx as a template.

Python docx aswell as docxtpl is used to run the program.

Files:
- format.xls is the Excel Spreadsheet populated with Dummy Data.
    - This Dummy Data showcases the use of Headers, Sub-Headers, Paragraphgs, and Tables.
- excel2doc_template.docx is the Word Document Template that the code uses to create the new Word Document.
    - multiple prefixes can be seen in this document such as {{document_title}}, and {{revision}}.
- new_test.docx is the newly created Word Document populated with the contents inside the Excel Spreadsheet.

Future Directions:
- I believe storing procedural documentation in an excel format can be beneficial. It allows companies to catagorize and search for data in a way that can add a tremendous amount of value. In a later version I hope to create a UI where users can interact with their own variables and provide their own data.
